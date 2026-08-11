import win32com.client
from pypdf import PdfReader, PdfWriter

import win32com.client


def convert_docx_to_pdf(docx_path: str, pdf_path: str):
    """Launches MS Word in the background to export a PDF and returns the app instance."""
    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False
    word_app.DisplayAlerts = 0

    doc_obj = word_app.Documents.Open(docx_path)
    doc_obj.ExportAsFixedFormat(
        OutputFileName=pdf_path,
        ExportFormat=17,  # wdExportFormatPDF
        OpenAfterExport=False,
        OptimizeFor=0,  # wdExportOptimizeForPrint
        CreateBookmarks=0  # wdExportCreateNoBookmarks
    )
    doc_obj.Close(False)
    return word_app  # Handed back up to main() for lifecycle management


def extract_sample_page(master_pdf_path: str, output_sample_path: str) -> None:
    """Isolates sample checking parameters down onto page index 0."""
    pdf_reader = PdfReader(master_pdf_path)
    pdf_writer = PdfWriter()

    pdf_writer.add_page(pdf_reader.pages[0])

    with open(output_sample_path, "wb") as sample_out:
        pdf_writer.write(sample_out)




