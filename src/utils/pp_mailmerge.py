import os
import pandas as pd
from mailmerge import MailMerge
import win32com.client  # Direct Windows Word Application Interface
from pypdf import PdfReader, PdfWriter

# Import the shared module engine safely from your directory layout
from barcode_generator import insert_code128_barcode

# 1. Setup explicit file configuration parameters
EXCEL_FILE = os.path.abspath("C:\\Users\\joseph.adogeri\\Desktop\\test\\pp.xls")
TEMPLATE_DOC = os.path.abspath("C:\\Users\\joseph.adogeri\\Desktop\\test\\pp.docx")
OUTPUT_DIR = os.path.abspath("C:\\Users\\joseph.adogeri\\Desktop\\test\\generated_letters")

# Ensure output directory exists safely
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Define the exact 3 files requested for output
MASTER_DOCX_PATH = os.path.normpath(os.path.join(OUTPUT_DIR, "test_parcel_parcelid.docx"))
MASTER_PDF_PATH = os.path.normpath(os.path.join(OUTPUT_DIR, "test_parcel_parcelid.pdf"))
TEST_SAMPLE_PDF_PATH = os.path.normpath(os.path.join(OUTPUT_DIR, "test_file.pdf"))

word_app = None

try:
    # 2. Read legacy format .xls file
    df = pd.read_excel(EXCEL_FILE, engine="xlrd", header=0)
    df.columns = df.columns.str.strip()

    print(f"Successfully loaded '{EXCEL_FILE}' containing {len(df)} rows.")
    print("Compiling global row database records...")

    # 3. Standard Data Cleaning and Configuration Pipeline
    cleaned_records_list = []

    # NEW TRACKER: List array capturing all unique parcel IDs matching the XLS order
    all_parcel_ids_list = []

    for index, row in df.iterrows():
        record = row.to_dict()
        clean_record = {}

        for key, val in record.items():
            key_str = str(key).strip()

            if pd.isna(val):
                clean_record[key_str] = ""
            elif isinstance(val, (int, float)) and key_str in ['TAX_YEAR', 'BATCH_NO', 'BATCH_ITEM_NO']:
                clean_record[key_str] = str(int(val))
            elif isinstance(val, (int, float)):
                if val % 1 == 0:
                    clean_record[key_str] = f"{int(val):,}"
                else:
                    clean_record[key_str] = f"{val:,.2f}"
            else:
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
    with MailMerge(TEMPLATE_DOC) as document:
        document.merge_templates(cleaned_records_list, separator="page_break")
        document.write(MASTER_DOCX_PATH)
    print(f"  > Saved Master Word document to: {MASTER_DOCX_PATH}")

    # =========================================================================
    # STEP B: INJECT UNIQUE CODE 128 BARCODES SEPARATELY FOR EVERY RECIPIENT
    # =========================================================================
    print("Processing global placeholder replacements with unique dynamic barcodes...")
    # UPDATED: Passing the entire ordered list array here instead of just one ID string
    insert_code128_barcode(MASTER_DOCX_PATH, MASTER_DOCX_PATH, all_parcel_ids_list, OUTPUT_DIR)

    # =========================================================================
    # STEP C: CONVERT THE COMPLETE MASTER FILE INTO A SINGLE MASTER PDF
    # =========================================================================
    print("Spawning automated Microsoft Word application layer for PDF generation...")
    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False
    word_app.DisplayAlerts = 0

    doc_obj = word_app.Documents.Open(MASTER_DOCX_PATH)
    doc_obj.ExportAsFixedFormat(
        OutputFileName=MASTER_PDF_PATH,
        ExportFormat=17,  # wdExportFormatPDF constant
        OpenAfterExport=False,
        OptimizeFor=0,  # wdExportOptimizeForPrint
        CreateBookmarks=0  # wdExportCreateNoBookmarks
    )
    doc_obj.Close(False)
    print(f"  > Saved Master PDF document layer to: {MASTER_PDF_PATH}")

    # =========================================================================
    # STEP D: EXTRACT FIRST PAGE TO GENERATE THE SINGLE TEST FILE
    # =========================================================================
    print("Extracting first record page sequence to construct sample validation file...")
    pdf_reader = PdfReader(MASTER_PDF_PATH)
    pdf_writer = PdfWriter()

    # Append the very first page object safely
    pdf_writer.add_page(pdf_reader.pages[0])

    with open(TEST_SAMPLE_PDF_PATH, "wb") as sample_out:
        pdf_writer.write(sample_out)
    print(f"  > Saved Single Page Validation File to: {TEST_SAMPLE_PDF_PATH}")

    print("\n✅ Success! Each letter inside the master document has been stamped with its own unique barcode.")

except FileNotFoundError:
    print(f"❌ Error: Missing assets. Ensure paths to Excel and Word are correct.")
except Exception as e:
    print(f"❌ A fatal runtime disruption occurred: {e}")

finally:
    if word_app is not None:
        try:
            word_app.Quit()
        except Exception:
            pass
