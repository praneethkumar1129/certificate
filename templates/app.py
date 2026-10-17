import pymupdf

doc = pymupdf.open("participation_template.pdf")

for font in doc.get_page_fonts(0, full=True):
    xref, ext, font_type, basefont, resource = font[:5]

    if "XBNiloofar" in basefont:
        name, extension, font_type, content = doc.extract_font(xref)

        with open(f"certificate_font.{extension}", "wb") as f:
            f.write(content)

        print(f"Extracted: certificate_font.{extension}")
        break