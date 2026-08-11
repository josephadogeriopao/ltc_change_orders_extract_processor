import os
import json
import pandas as pd
from mailmerge import MailMerge
import win32com.client
from pypdf import PdfReader, PdfWriter
from .barcode_generator import insert_code128_barcode


def process_change_orders(excel_file, template_doc, output_dir, master_docx_path, master_pdf_path,
                          test_sample_pdf_path):
    """
    Processes JSON or Excel records, merges them into a template, inserts barcodes,
    and converts the final outputs into Master and Sample PDF documents with strict
    integer comma formatting (no decimals).
    """
    word_app = None

    # 1. Ensure output directory exists safely
    os.makedirs(output_dir, exist_ok=True)

    # 2. Dynamic Source Engine Detection (.json vs .xls vs .xlsx)
    if excel_file.lower().endswith(".json"):
        with open(excel_file, "r", encoding="utf-8") as f:
            raw_records = json.load(f)
        if isinstance(raw_records, dict):
            raw_records = [raw_records]
        print(f"Successfully loaded '{excel_file}' containing {len(raw_records)} JSON rows.")
    else:
        chosen_engine = "openpyxl" if excel_file.lower().endswith(".xlsx") else "xlrd"
        df = pd.read_excel(excel_file, engine=chosen_engine, header=0)
        df.columns = df.columns.str.strip()
        raw_records = df.to_dict(orient="records")
        print(f"Successfully loaded '{excel_file}' using engine '{chosen_engine}' containing {len(df)} rows.")

    print("Compiling global row database records...")

    # 3. Standard Data Cleaning and Configuration Pipeline
    cleaned_records_list = []
    all_parcel_ids_list: list[str] = []

    # Explicit sets for strict control of data field models
    numeric_id_fields = {'TAX_YEAR', 'BATCH_NO', 'BATCH_ITEM_NO'}

    for record in raw_records:
        clean_record = {}

        for key, val in record.items():
            key_str = str(key).strip()

            # Fix pd.isna fallback check for dict environments safely
            is_val_nan = pd.isna(val) if 'pd' in globals() and hasattr(pd, "isna") else (val is None or val == "")

            if is_val_nan:
                clean_record[key_str] = ""
            elif isinstance(val, (int, float)) and key_str in numeric_id_fields:
                clean_record[key_str] = str(int(val))
            elif isinstance(val, (int, float)):
                # Decisively strip decimals (.00) and format with commas for thousands
                if val % 1 == 0:
                    clean_record[key_str] = f"{int(val):,}"
                else:
                    # Fallback configuration for real floats, stripping trailing .00 if it evaluates flat
                    formatted_val = f"{val:,.2f}"
                    if formatted_val.endswith(".00"):
                        clean_record[key_str] = formatted_val[:-3]
                    else:
                        clean_record[key_str] = formatted_val
            else:
                # Clean text layout variants and remove heavy spacing noise
                clean_record[key_str] = str(val).strip()

        # Case variant bindings for text layout targets
        reason_value = clean_record.get("REASON", "")
        clean_record["REASON"] = reason_value
        clean_record["reason"] = reason_value
        clean_record["Reason"] = reason_value

        dynamic_parcel_id = clean_record.get("PARCELID", "").strip()
        clean_record['parcelid'] = dynamic_parcel_id
        clean_record['PARCELID'] = dynamic_parcel_id

        # Save this unique value to the sequential array index list
        all_parcel_ids_list.append(dynamic_parcel_id)

        # Address mapping rules
        addr2_val = clean_record.get("TAXPAYER_ADDR2", "")
        addr3_val = clean_record.get("TAXPAYER_ADDR3", "")
        if addr2_val == "":
            clean_record["TAXPAYER_ADDR2"] = addr3_val
            clean_record["TAXPAYER_ADDR3"] = ""

        # Remove timestamp noise from date fields safely
        for date_col in ['STATUS_DATE', 'BATCH_SUBMITTED']:
            if clean_record.get(date_col) and " " in clean_record[date_col]:
                clean_record[date_col] = clean_record[date_col].split(" ")[0]

        cleaned_records_list.append(clean_record)

    # =========================================================================
    # STEP A: MERGE ALL ROWS INTO A SINGLE UNIFIED MASTER DOCX FILE
    # =========================================================================
    print("\nMerging all template data rows sequentially into one document...")
    with MailMerge(template_doc) as document:
        document.merge_templates(cleaned_records_list, separator="page_break")
        document.write(master_docx_path)
    print(f"  > Saved Master Word document to: {master_docx_path}")

    # =========================================================================
    # STEP B: INJECT UNIQUE CODE 128 BARCODES SEPARATELY FOR EVERY RECIPIENT
    # =========================================================================
    print("Processing global placeholder replacements with unique dynamic barcodes...")
    insert_code128_barcode(master_docx_path, master_docx_path, all_parcel_ids_list, output_dir)

    # =========================================================================
    # STEP C: CONVERT THE COMPLETE MASTER FILE INTO A SINGLE MASTER PDF
    # =========================================================================
    print("Spawning automated Microsoft Word application layer for PDF generation...")
    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False
    word_app.DisplayAlerts = 0

    doc_obj = word_app.Documents.Open(master_docx_path)
    doc_obj.ExportAsFixedFormat(
        OutputFileName=master_pdf_path,
        ExportFormat=17,  # wdExportFormatPDF constant
        OpenAfterExport=False,
        OptimizeFor=0,  # wdExportOptimizeForPrint
        CreateBookmarks=0  # wdExportCreateNoBookmarks
    )
    doc_obj.Close(False)
    print(f"  > Saved Master PDF document layer to: {master_pdf_path}")

    # =========================================================================
    # STEP D: EXTRACT FIRST PAGE TO GENERATE THE SINGLE TEST FILE
    # =========================================================================
    print("Extracting first record page sequence to construct sample validation file...")
    pdf_reader = PdfReader(master_pdf_path)
    pdf_writer = PdfWriter()

    # Append the very first page object safely
    pdf_writer.add_page(pdf_reader.pages[0])

    with open(test_sample_pdf_path, "wb") as sample_out:
        pdf_writer.write(sample_out)
    print(f"  > Saved Single Page Validation File to: {test_sample_pdf_path}")

    return word_app
