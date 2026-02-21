import win32com.client
from pathlib import Path
from barcode_generator import generate_barcode


def generate_all_labels(
    template_path,
    excel_path,   # unused but kept for compatibility
    sheet_name,   # unused but kept for compatibility
    records,
    out_docx_dir,
    out_pdf_dir
):
    out_docx_dir = Path(out_docx_dir)
    out_pdf_dir = Path(out_pdf_dir)
    out_docx_dir.mkdir(parents=True, exist_ok=True)
    out_pdf_dir.mkdir(parents=True, exist_ok=True)

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        # 🔹 OPEN TEMPLATE (Already linked)
        doc = word.Documents.Open(str(Path(template_path).resolve()))
        
        # 🔥 ATTACH EXCEL ONLY ONCE
        doc.MailMerge.OpenDataSource(
            Name=str(Path(excel_path).resolve()),
            ConfirmConversions=False,
            ReadOnly=True,
            LinkToSource=False,
            AddToRecentFiles=False,
            SQLStatement="SELECT * FROM [Sheet1$]"
        )


        for idx, record in enumerate(records, start=1):

            ean_value = record["EAN_No"]

            # Handle duplicate filenames
            base_docx = out_docx_dir / f"{ean_value}.docx"
            base_pdf = out_pdf_dir / f"{ean_value}.pdf"

            out_docx = base_docx
            out_pdf = base_pdf
            counter = 1

            while out_docx.exists() or out_pdf.exists():
                out_docx = out_docx_dir / f"{ean_value}_{counter}.docx"
                out_pdf = out_pdf_dir / f"{ean_value}_{counter}.pdf"
                counter += 1

            # 🔹 Lock to single record
            doc.MailMerge.DataSource.FirstRecord = idx
            doc.MailMerge.DataSource.LastRecord = idx
            doc.MailMerge.DataSource.ActiveRecord = idx

            doc.MailMerge.Execute(False)
            merged = word.ActiveDocument

            # 🔹 Insert Barcode
            barcode_img = generate_barcode(ean_value, Path("output/resources"))
            barcode_img = str(Path(barcode_img).resolve())

            find_barcode = merged.Content.Find
            find_barcode.Text = "<<BARCODE>>"
            find_barcode.Wrap = 1

            while find_barcode.Execute():
                rng = find_barcode.Parent
                rng.Text = ""
                pic = rng.InlineShapes.AddPicture(
                    FileName=barcode_img,
                    LinkToFile=False,
                    SaveWithDocument=True
                )
                pic.LockAspectRatio = True
                pic.Width = 130

            # 🔹 Insert EAN below barcode
            find_ean = merged.Content.Find
            find_ean.Text = "<<EAN>>"
            find_ean.Wrap = 1

            while find_ean.Execute():
                para = find_ean.Parent
                para.Text = ean_value
                para.Font.Name = "Times New Roman"
                para.Font.Size = 9
                para.ParagraphFormat.Alignment = 1

            merged.SaveAs(str(out_docx.resolve()))
            merged.SaveAs(str(out_pdf.resolve()), FileFormat=17)

            merged.Close(False)

            print(f"✔ Generated labels for EAN: {ean_value}")

        doc.Close(False)

    finally:
        word.Quit()