from PyPDF2 import PdfMerger
import os
import docx2pdf
import argparse

def convert_docx_to_pdf(docx_file, pdf_file):
    try:
        docx2pdf.convert(docx_file, pdf_file)
        print(f"Successfully converted {docx_file} to {pdf_file}")
    except Exception as e:
        print(f"Error converting {docx_file}: {e}")

def get_file_names(folder_path):
    file_names = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            file_names.append(file)
    return file_names

def main_tool(first_pdf, docx_folder):
    pdf_folder = f"{docx_folder}_pdfs"
    if not os.path.exists(pdf_folder):
        os.mkdir(pdf_folder)
    out_folder = f"{docx_folder}_edited"
    if not os.path.exists(out_folder):
        os.mkdir(out_folder)

    for filename in os.listdir(docx_folder):
        if filename.endswith(".docx"):
            docx_path = os.path.join(docx_folder, filename)
            pdf_path = os.path.join(pdf_folder, filename[:-5] + ".pdf")
            convert_docx_to_pdf(docx_path, pdf_path)

    files = get_file_names(pdf_folder)
    for i in range(0, len(files)):
        merger = PdfMerger()
        pdfs = [first_pdf, os.path.join(pdf_folder, files[i])]
        for pdf in pdfs:
            merger.append(pdf)
        merger.write(f"{out_folder}/{files[i]}")
        merger.close()

    print("Successfully Converted and Merger!!!!")

def main():
    parser = argparse.ArgumentParser(description="Tool for MK Mam to Merge a first PDF with DOCX files converted to PDF given in specific folder")
    parser.add_argument(
        "-file", 
        type=str, 
        required=True,
        help="Path to the first PDF file that will be merged at top of each doc"
    )
    parser.add_argument(
        "-folder", 
        type=str, 
        required=True,
        help="Path to the folder containing DOCX files to be converted and merged."
    )

    args = parser.parse_args()
    main_tool(args.file, args.folder)

if __name__ == "__main__":
    main()

# if __name__ == "__main__":
#     first_pdf = "first.pdf"
#     docx_folder = "some_docx"
#     main_tool(first_pdf, docx_folder)