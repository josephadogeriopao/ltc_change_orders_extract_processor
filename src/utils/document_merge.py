from mailmerge import MailMerge
from .barcode_generator import insert_code128_barcode

def generate_master_docx(template_doc: str, cleaned_records: list[dict], output_path: str) -> None:
    """Performs Word MailMerge operation across database rows."""
    with MailMerge(template_doc) as document:
        document.merge_templates(cleaned_records, separator="page_break")
        document.write(output_path)

def inject_barcodes(docx_path: str, parcel_ids: list[str], output_dir: str) -> None:
    """Wraps external generator injection pattern."""
    insert_code128_barcode(docx_path, docx_path, parcel_ids, output_dir)
