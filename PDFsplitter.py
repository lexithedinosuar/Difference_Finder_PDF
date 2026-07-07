from pypdf import PdfReader, PdfWriter
from pathlib import Path
import os


def PDFsplit (FileName, FolderName):
    """
    takes a PDF document (FileName) and splits it, seperating each page into it's own document
    and adding each document document into a new folder with a given FolderName.

    returns the file location of the output folder
    """
    output_folder = Path(FolderName)
    output_folder.mkdir(exist_ok=True)

    reader = PdfReader(FileName)

    for pn, page in enumerate(reader.pages, start= 1):
        writer = PdfWriter()
        writer.add_page(page)

        output_file = output_folder / f"page_{pn}.pdf"

        with open(output_file, "wb") as f:
            writer.write(f)

    return output_folder

def FullPDF_to_SplitTXT(pdf_file, folder_name):
    # set up output folder
    output_folder = Path(folder_name)
    output_folder.mkdir(exist_ok=True)

    # read PDF
    reader = PdfReader(pdf_file)

    # extract all text
    for pn, page in enumerate(reader.pages, start=1):
        text=""
        text += (page.extract_text() or "") + "\n"
        txt_file = output_folder / f"page_{pn}.txt"

        with open(txt_file, "w", encoding="utf-8") as f:
            f.write(text)
    return output_folder


def read_txt_file(txt_file):
    with open(txt_file, "r", encoding="utf-8") as f:
        text = f.read()

    return text
