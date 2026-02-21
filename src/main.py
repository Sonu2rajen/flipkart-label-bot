import sys
from pathlib import Path
from excel_reader import read_excel
from word_com_filler import generate_all_labels


def get_base_path():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent


def main():
    print("\n=== FLIPKART LABEL BOT ===\n")

    BASE_DIR = get_base_path()

    excel_file = input("Enter Excel file name (example: Flipkart.xlsx): ").strip()
    sheet_name = input("Enter Sheet name (example: Sheet1): ").strip()

    excel_path = BASE_DIR / excel_file
    template_path = BASE_DIR / "template" / "Flipkart.docx"
    out_docx_dir = BASE_DIR / "output" / "docx"
    out_pdf_dir = BASE_DIR / "output" / "pdf"

    try:
        records = read_excel(excel_path, sheet_name)

        if not records:
            print("No valid records found.")
            return

        generate_all_labels(
        template_path=template_path,
        excel_path=excel_path,
        sheet_name=sheet_name,
        records=records,
        out_docx_dir=out_docx_dir,
        out_pdf_dir=out_pdf_dir
    )

        print("\n✔ All labels generated successfully.\n")

    except Exception as e:
        print("\n❌ FULL ERROR:")
        import traceback
        traceback.print_exc()

    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
