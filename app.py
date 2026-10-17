
import os
import re
import io
import csv
import json
import uuid
import smtplib
import zipfile
from pathlib import Path
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr

import pandas as pd
import fitz  # PyMuPDF
from dotenv import load_dotenv
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = BASE_DIR / "Certificate_AURA26_with_signatures.pdf"
JOBS_DIR = BASE_DIR / "jobs"
JOBS_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Participation Certificate Generator")
app.mount(
    "/static",
    StaticFiles(directory=str(BASE_DIR / "static")),
    name="static",
)

JOBS = {}


# ---------------------------------------------------------
# EXCEL / CSV HELPERS
# ---------------------------------------------------------

def clean_cell(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_header(value):
    return re.sub(r"\s+", " ", str(value).strip()).lower()


def read_registration_file(filename: str, content: bytes) -> pd.DataFrame:
    suffix = Path(filename).suffix.lower()

    try:
        if suffix in (".xlsx", ".xlsm", ".xls"):
            return pd.read_excel(
                io.BytesIO(content), dtype=str
            ).fillna("")

        if suffix == ".csv":
            try:
                return pd.read_csv(
                    io.BytesIO(content),
                    dtype=str,
                    encoding="utf-8-sig",
                ).fillna("")
            except UnicodeDecodeError:
                return pd.read_csv(
                    io.BytesIO(content),
                    dtype=str,
                    encoding="latin-1",
                ).fillna("")

    except Exception as exc:
        raise HTTPException(
            status_code=400,
            detail=f"Could not read spreadsheet: {exc}",
        )

    raise HTTPException(
        status_code=400,
        detail="Upload an .xlsx, .xlsm, .xls, or .csv file.",
    )


def find_column(columns, exact_names):
    normalized = {
        normalize_header(column): column
        for column in columns
    }

    for name in exact_names:
        key = normalize_header(name)
        if key in normalized:
            return normalized[key]

    return None


def extract_students(df: pd.DataFrame):
    students = []
    columns = list(df.columns)

    team_email_col = find_column(
        columns,
        ["Email Address", "Email", "Team Email Address"],
    )
    college_col = find_column(
        columns,
        [
            "College / Institution Name",
            "College Name",
            "Institution Name",
            "College",
        ],
    )
    team_col = find_column(columns, ["Team Name"])

    for row_index, (_, row) in enumerate(
        df.iterrows(), start=2
    ):
        team_email = (
            clean_cell(row[team_email_col])
            if team_email_col else ""
        )
        college = (
            clean_cell(row[college_col])
            if college_col else ""
        )
        team = (
            clean_cell(row[team_col])
            if team_col else ""
        )

        for member_num in range(1, 21):
            name_col = find_column(
                columns,
                [
                    f"Member {member_num} - Full Name",
                    f"Member {member_num} Full Name",
                    f"Member {member_num} Name",
                ],
            )

            email_col = find_column(
                columns,
                [
                    f"Member {member_num} - Email Address",
                    f"Member {member_num} Email Address",
                    f"Member {member_num} - Email",
                    f"Member {member_num} Email",
                ],
            )

            if not name_col and not email_col:
                if member_num > 4:
                    break
                continue

            name = (
                clean_cell(row[name_col])
                if name_col else ""
            )
            email = (
                clean_cell(row[email_col])
                if email_col else ""
            )

            if member_num == 1 and not email:
                email = team_email

            if not name and not email:
                continue

            issue = ""

            if not name:
                issue = "Missing student name"
            elif not email:
                issue = "Missing email address"
            elif not re.match(
                r"^[^@\s]+@[^@\s]+\.[^@\s]+$",
                email,
            ):
                issue = "Invalid email address"

            students.append({
                "name": name,
                "email": email,
                "college": college,
                "team": team,
                "source_row": row_index,
                "member": member_num,
                "issue": issue,
            })

    unique = []
    seen = set()

    for student in students:
        key = (
            student["name"].casefold(),
            student["email"].casefold(),
        )

        if key in seen:
            continue

        seen.add(key)
        unique.append(student)

    return unique


# ---------------------------------------------------------
# CERTIFICATE GENERATION
#
# Template: Certificate_AURA26_with_signatures.pdf (842.25 x 595.5 pt, landscape)
# Original fonts: TimesNewRomanPSMT (15 pt) for fill-in lines.
# Inserted font:  "tiro" (PyMuPDF built-in Times-Roman) — same family, visually
#                 indistinguishable at this size, and guaranteed to render all
#                 Latin characters correctly.
#
# XB Niloofar Bold (templates/certificate_font.ttf) is a Persian/Arabic font;
# it is missing most Latin glyphs (e.g. 'B', 'N') and must NOT be used here.
#
# Field positions (measured from the template PDF):
#   NAME    : x=240, baseline_y=290  (after "Mr./Ms.", underscore x=236.88..767.64)
#   COLLEGE : x=56,  baseline_y=331  (left underscore x=56.07..361.91)
#   EVENT   : x=553, baseline_y=331  (right underscore x=549.31..782.66)
# ---------------------------------------------------------

# Shared font instance for width calculations (module-level, created once).
_TIRO_FONT = fitz.Font(fontname="tiro")


def _insert_field(
    page: fitz.Page,
    text: str,
    x: float,
    baseline_y: float,
    max_width: float,
    nominal_size: float = 15.0,
    min_size: float = 8.0,
) -> None:
    """Insert *text* into *page* using Times-Roman, shrinking to fit max_width."""
    text = str(text or "").strip()
    if not text:
        return

    size = nominal_size
    while size > min_size:
        if _TIRO_FONT.text_length(text, fontsize=size) <= max_width:
            break
        size = max(min_size, size - 0.5)

    page.insert_text(
        (x, baseline_y),
        text,
        fontname="tiro",
        fontsize=size,
        color=(0, 0, 0),
        overlay=True,
    )


def create_certificate(
    template_bytes: bytes,
    student: dict,
    event_name: str,
) -> bytes:
    try:
        doc = fitz.open(stream=template_bytes, filetype="pdf")
    except Exception as exc:
        raise ValueError(f"Could not open certificate PDF: {exc}")

    if len(doc) < 1:
        doc.close()
        raise ValueError("The template PDF has no pages.")

    try:
        page = doc[0]

        name = str(student.get("name", "")).strip()
        college = str(student.get("college", "")).strip()
        event_name = str(event_name or "").strip()

        # NAME — after "This is to certify that Mr./Ms."
        # Safe width: from x=240 to just before " of " at x=765.39  →  524 pt
        _insert_field(page, name, x=240, baseline_y=290, max_width=524)

        # COLLEGE — left blank before "has actively participated in the"
        # Safe width: from x=56 to x=361.91  →  302 pt
        if college:
            _insert_field(page, college, x=56, baseline_y=331, max_width=302)

        # EVENT — right blank after "has actively participated in the"
        # Safe width: from x=553 to x=782.66  →  229 pt
        _insert_field(page, event_name, x=553, baseline_y=331, max_width=229)

        output = io.BytesIO()
        doc.save(output, garbage=4, deflate=True)
        return output.getvalue()

    finally:
        doc.close()


def safe_filename(value):
    value = re.sub(
        r"[^A-Za-z0-9._-]+",
        "_",
        value.strip(),
    )
    return value.strip("._")[:80] or "Student"


# ---------------------------------------------------------
# EMAIL CONFIGURATION
# ---------------------------------------------------------

def smtp_settings():
    return {
        "host": os.getenv(
            "SMTP_HOST", "smtp.gmail.com"
        ),
        "port": int(
            os.getenv("SMTP_PORT", "587")
        ),
        "user": os.getenv(
            "SMTP_USER", ""
        ).strip(),
        "password": os.getenv(
            "SMTP_APP_PASSWORD", ""
        ).replace(" ", "").strip(),
        "sender_name": os.getenv(
            "SENDER_NAME", "Event Certificates"
        ).strip(),
    }


def send_certificate_email(
    settings,
    recipient,
    event_name,
    pdf_bytes,
    filename,
):
    if not settings["user"] or not settings["password"]:
        raise RuntimeError(
            "Gmail is not configured. Set SMTP_USER "
            "and SMTP_APP_PASSWORD in .env."
        )

    message = EmailMessage()
    message["Subject"] = (
        f"Your Participation Certificate – {event_name}"
    )
    message["From"] = formataddr(
        (
            settings["sender_name"],
            settings["user"],
        )
    )
    message["To"] = recipient

    message.set_content(
        f"Hello,\n\n"
        f"Thank you for participating in {event_name}.\n"
        "Please find your participation certificate "
        "attached as a PDF.\n\n"
        "Regards,\nEvent Organizing Team"
    )

    message.add_attachment(
        pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=filename,
    )

    with smtplib.SMTP(
        settings["host"],
        settings["port"],
        timeout=30,
    ) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(
            settings["user"],
            settings["password"],
        )
        server.send_message(message)


# ---------------------------------------------------------
# INPUT LOADING
# ---------------------------------------------------------

async def load_inputs(
    excel: UploadFile,
    template: UploadFile | None,
):
    excel_bytes = await excel.read()

    df = read_registration_file(
        excel.filename or "registrations.xlsx",
        excel_bytes,
    )

    if template and template.filename:
        template_bytes = await template.read()

        if not template_bytes.startswith(b"%PDF"):
            raise HTTPException(
                status_code=400,
                detail="Certificate template must be a PDF.",
            )
    else:
        if not DEFAULT_TEMPLATE.exists():
            raise HTTPException(
                status_code=500,
                detail="Default certificate template is missing.",
            )

        template_bytes = DEFAULT_TEMPLATE.read_bytes()

    students = extract_students(df)

    if not students:
        raise HTTPException(
            status_code=400,
            detail=(
                "No students found. Check that your sheet has "
                "'Member 1 - Full Name' and member email columns."
            ),
        )

    return students, template_bytes


# ---------------------------------------------------------
# WEB ROUTES
# ---------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def home():
    return (
        BASE_DIR / "static" / "index.html"
    ).read_text(encoding="utf-8")


@app.get("/api/health")
async def health():
    return {"status": "ok"}


@app.post("/api/preview")
async def preview(
    excel: UploadFile = File(...),
    event_name: str = Form(...),
    template: UploadFile | None = File(None),
):
    event_name = event_name.strip()

    if not event_name:
        raise HTTPException(
            status_code=400,
            detail="Enter an event name.",
        )

    students, template_bytes = await load_inputs(
        excel, template
    )

    eligible = [
        student for student in students
        if not student.get("issue")
        and student.get("name")
        and student.get("email")
    ]

    if not eligible:
        raise HTTPException(
            status_code=400,
            detail=(
                "No valid student with both name and email "
                "was found."
            ),
        )

    pdf_bytes = create_certificate(
        template_bytes,
        eligible[0],
        event_name,
    )

    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={
            "Content-Disposition":
                'inline; filename="certificate-preview.pdf"'
        },
    )


@app.post("/api/recipients")
async def preview_recipients(
    excel: UploadFile = File(...),
    event_name: str = Form(...),
    template: UploadFile | None = File(None),
):
    if not event_name.strip():
        raise HTTPException(
            status_code=400,
            detail="Enter an event name.",
        )

    students, _template_bytes = await load_inputs(
        excel, template
    )

    return {
        "event_name": event_name.strip(),
        "total": len(students),
        "eligible": sum(
            1 for student in students
            if not student.get("issue")
            and student.get("name")
            and student.get("email")
        ),
        "recipients": [
            {
                "name": student.get("name", ""),
                "email": student.get("email", ""),
                "college": student.get("college", ""),
                "team": student.get("team", ""),
                "issue": student.get("issue", ""),
            }
            for student in students
        ],
    }


@app.post("/api/send")
async def send_all(
    excel: UploadFile = File(...),
    event_name: str = Form(...),
    confirm_send: str = Form(...),
    template: UploadFile | None = File(None),
):
    event_name = event_name.strip()

    if not event_name:
        raise HTTPException(
            status_code=400,
            detail="Enter an event name.",
        )

    if confirm_send.lower() not in (
        "true", "yes", "1", "on"
    ):
        raise HTTPException(
            status_code=400,
            detail="Confirm bulk sending before continuing.",
        )

    students, template_bytes = await load_inputs(
        excel, template
    )

    job_id = uuid.uuid4().hex
    job_dir = JOBS_DIR / job_id
    cert_dir = job_dir / "certificates"
    cert_dir.mkdir(parents=True, exist_ok=True)

    settings = smtp_settings()
    report = []
    zip_path = job_dir / "certificates.zip"

    for student in students:
        name = student["name"]
        email = student["email"]

        if student.get("issue"):
            report.append({
                **student,
                "status": "Skipped",
                "message": student["issue"],
                "filename": "",
            })
            continue

        if not name or not email:
            report.append({
                **student,
                "status": "Skipped",
                "message": "Missing name or email",
                "filename": "",
            })
            continue

        filename = (
            f"{safe_filename(name)}_certificate.pdf"
        )

        try:
            pdf_bytes = create_certificate(
                template_bytes,
                student,
                event_name,
            )

            (
                cert_dir / filename
            ).write_bytes(pdf_bytes)

            try:
                send_certificate_email(
                    settings,
                    email,
                    event_name,
                    pdf_bytes,
                    filename,
                )

                status = "Sent"
                message = "Email accepted by SMTP server"

            except Exception as exc:
                raw_error = str(exc)

                if (
                    isinstance(exc, OSError)
                    and getattr(exc, "winerror", None) == 11001
                ):
                    raw_error = (
                        "DNS lookup failed for the SMTP server. "
                        "Verify SMTP_HOST=smtp.gmail.com and "
                        "SMTP_PORT=587 in .env. Run "
                        "Resolve-DnsName smtp.gmail.com in "
                        "PowerShell. Original error: "
                        f"{raw_error}"
                    )

                status = "Failed"
                message = raw_error[:500]

            report.append({
                **student,
                "status": status,
                "message": message,
                "filename": filename,
            })

        except Exception as exc:
            report.append({
                **student,
                "status": "Failed",
                "message": (
                    f"Certificate generation failed: {exc}"
                )[:500],
                "filename": "",
            })

    with zipfile.ZipFile(
        zip_path,
        "w",
        compression=zipfile.ZIP_DEFLATED,
    ) as archive:
        for pdf in cert_dir.glob("*.pdf"):
            archive.write(pdf, arcname=pdf.name)

    report_path = job_dir / "report.csv"

    report_df = pd.DataFrame([
        {
            "Student Name": item.get("name", ""),
            "Email": item.get("email", ""),
            "College": item.get("college", ""),
            "Team": item.get("team", ""),
            "Source Excel Row": item.get("source_row", ""),
            "Member Number": item.get("member", ""),
            "Status": item.get("status", ""),
            "Message": item.get("message", ""),
            "Certificate File": item.get("filename", ""),
        }
        for item in report
    ])

    report_df.to_csv(report_path, index=False)

    job = {
        "job_id": job_id,
        "event_name": event_name,
        "created_at": datetime.now().isoformat(
            timespec="seconds"
        ),
        "total": len(report),
        "sent": sum(
            1 for item in report
            if item["status"] == "Sent"
        ),
        "failed": sum(
            1 for item in report
            if item["status"] == "Failed"
        ),
        "skipped": sum(
            1 for item in report
            if item["status"] == "Skipped"
        ),
        "report": report,
        "zip_path": str(zip_path),
        "report_path": str(report_path),
    }

    JOBS[job_id] = job

    return JSONResponse({
        "job_id": job_id,
        "event_name": event_name,
        "created_at": job["created_at"],
        "total": job["total"],
        "sent": job["sent"],
        "failed": job["failed"],
        "skipped": job["skipped"],
        "report": [
            {
                "name": item.get("name", ""),
                "email": item.get("email", ""),
                "status": item.get("status", ""),
                "message": item.get("message", ""),
                "filename": item.get("filename", ""),
            }
            for item in report
        ],
        "download_url": (
            f"/api/jobs/{job_id}/download"
        ),
        "report_url": (
            f"/api/jobs/{job_id}/report"
        ),
    })


@app.get("/api/jobs/{job_id}/download")
async def download_zip(job_id: str):
    job = JOBS.get(job_id)

    if (
        not job
        or not Path(job["zip_path"]).exists()
    ):
        raise HTTPException(
            status_code=404,
            detail=(
                "Job not found. Jobs are cleared when "
                "the server restarts."
            ),
        )

    return FileResponse(
        job["zip_path"],
        filename=f"certificates_{job_id[:8]}.zip",
        media_type="application/zip",
    )


@app.get("/api/jobs/{job_id}/report")
async def download_report(job_id: str):
    job = JOBS.get(job_id)

    if (
        not job
        or not Path(job["report_path"]).exists()
    ):
        raise HTTPException(
            status_code=404,
            detail=(
                "Report not found. Jobs are cleared when "
                "the server restarts."
            ),
        )

    return FileResponse(
        job["report_path"],
        filename=f"email_report_{job_id[:8]}.csv",
        media_type="text/csv",
    )


@app.get("/api/jobs/{job_id}")
async def job_status(job_id: str):
    job = JOBS.get(job_id)

    if not job:
        raise HTTPException(
            status_code=404,
            detail=(
                "Job not found. Jobs are cleared when "
                "the server restarts."
            ),
        )

    return {
        "job_id": job["job_id"],
        "event_name": job["event_name"],
        "created_at": job["created_at"],
        "total": job["total"],
        "sent": job["sent"],
        "failed": job["failed"],
        "skipped": job["skipped"],
    }