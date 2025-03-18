import os
import comtypes.client

def convert_all_docx_to_pdf():
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False  # Run in background

    for file in os.listdir():
        if file.lower().endswith(".docx"):
            pdf_file = file.replace(".docx", ".pdf")
            input_path = os.path.abspath(file)
            output_path = os.path.abspath(pdf_file)

            try:
                doc = word.Documents.Open(input_path)
                doc.SaveAs(output_path, FileFormat=17)  # 17 is PDF format
                doc.Close()
                print(f"Converted: {pdf_file}")
            except Exception as e:
                print(f"Error converting {file}: {e}")

    word.Quit()

if __name__ == "__main__":
    convert_all_docx_to_pdf()

# pyinstaller --onefile --noconsole docx_to_pdf.py
