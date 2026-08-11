import os
import sys
from dotenv import load_dotenv
from utils.mailmerge import process_change_orders

def main():
    # 1. Automatically load configurations into environment workspace memory
    load_dotenv()

    # 2. Extract configuration tokens with programmatic fallback protections
    data_file_raw = os.getenv("DATA_FILE")
    template_doc_raw = os.getenv("TEMPLATE_DOC")
    output_dir_raw = os.getenv("OUTPUT_DIR")

    master_docx_name = os.getenv("MASTER_DOCX_NAME")
    master_pdf_name = os.getenv("MASTER_PDF_NAME")
    test_sample_pdf_name = os.getenv("TEST_SAMPLE_PDF_NAME")

    # Safe structural assertion step ensuring no key metrics are missing
    if not all([data_file_raw, template_doc_raw, output_dir_raw, master_docx_name, master_pdf_name, test_sample_pdf_name]):
        print("❌ Critical System Launch Fault: Active variables mapping is incomplete in .env configuration layout.")
        sys.exit(1)

    # 3. Handle explicit system paths resolution
    data_file = os.path.abspath(data_file_raw)
    template_doc = os.path.abspath(template_doc_raw)
    output_dir = os.path.abspath(output_dir_raw)

    master_docx_path = os.path.normpath(os.path.join(output_dir, master_docx_name))
    master_pdf_path = os.path.normpath(os.path.join(output_dir, master_pdf_name))
    test_sample_pdf_path = os.path.normpath(os.path.join(output_dir, test_sample_pdf_name))

    print("=" * 80)
    print(f"🚀 Initializing Change Order Framework Run Instance")
    print(f"   > Data Source Target:  {data_file}")
    print(f"   > Output Destination: {output_dir}")
    print("=" * 80)

    word_app = None

    try:
        # 4. Trigger processing framework engine exactly once
        word_app = process_change_orders(
            excel_file=data_file,
            template_doc=template_doc,
            output_dir=output_dir,
            master_docx_path=master_docx_path,
            master_pdf_path=master_pdf_path,
            test_sample_pdf_path=test_sample_pdf_path
        )

        print("\n✅ Success! Pipeline completed layout mapping tasks successfully.")

    except FileNotFoundError:
        print("\n❌ Error: Processing sequence dropped due to missing data file assets.")
    except Exception as e:
        print(f"\n❌ A fatal runtime disruption occurred during processing loop: {e}")

    finally:
        # 5. Clean up Word COM automation state safely
        if word_app is not None:
            try:
                word_app.Quit()
                print("🔒 Word Application background wrapper shutdown complete.")
            except Exception:
                pass

if __name__ == "__main__":
    main()


# import os
# from utils.mailmerge import process_change_orders
#
# '''
#     Processing Personal Property Change Orders
# '''
# # 1. Setup explicit file configuration parameters
# EXCEL_FILE = os.path.abspath("C:\\Users\\joseph.adogeri\\Desktop\\test\\pp.xls")
# TEMPLATE_DOC = os.path.abspath("../assets/templates/pp.docx")
# OUTPUT_DIR = os.path.abspath("C:\\Users\\joseph.adogeri\\Desktop\\test\\pp_generated_letters")
#
# # Define the exact 3 files requested for output
# MASTER_DOCX_PATH = os.path.normpath(os.path.join(OUTPUT_DIR, "pp_test_parcel_parcelid.docx"))
# MASTER_PDF_PATH = os.path.normpath(os.path.join(OUTPUT_DIR, "pp_test_parcel_parcelid.pdf"))
# TEST_SAMPLE_PDF_PATH = os.path.normpath(os.path.join(OUTPUT_DIR, "pp_test_file.pdf"))
#
# word_app = None
#
# try:
#     # Run the extracted engine function
#     word_app = process_change_orders(
#         excel_file=EXCEL_FILE,
#         template_doc=TEMPLATE_DOC,
#         output_dir=OUTPUT_DIR,
#         master_docx_path=MASTER_DOCX_PATH,
#         master_pdf_path=MASTER_PDF_PATH,
#         test_sample_pdf_path=TEST_SAMPLE_PDF_PATH
#     )
#
#     print("\n✅ Success! Each letter inside the master document has been stamped with its own unique barcode.")
#
# except FileNotFoundError:
#     print(f"❌ Error: Missing assets. Ensure paths to Excel and Word are correct.")
# except Exception as e:
#     print(f"❌ A fatal runtime disruption occurred: {e}")
#
# finally:
#     # Clean up Word COM automation state safely
#     if word_app is not None:
#         try:
#             word_app.Quit()
#             print("Word Application shutdown complete.")
#         except Exception:
#             pass
#
# '''
#     Processing Real Property Change Orders
# '''
# # 1. Setup explicit file configuration parameters
# EXCEL_FILE = os.path.abspath("C:\\Users\\joseph.adogeri\\Desktop\\test\\real.xls")
# TEMPLATE_DOC = os.path.abspath("../assets/templates/real.docx")
# OUTPUT_DIR = os.path.abspath("C:\\Users\\joseph.adogeri\\Desktop\\test\\real_generated_letters")
#
# # Define the exact 3 files requested for output
# MASTER_DOCX_PATH = os.path.normpath(os.path.join(OUTPUT_DIR, "real_test_parcel_parcelid.docx"))
# MASTER_PDF_PATH = os.path.normpath(os.path.join(OUTPUT_DIR, "real_test_parcel_parcelid.pdf"))
# TEST_SAMPLE_PDF_PATH = os.path.normpath(os.path.join(OUTPUT_DIR, "real_test_file.pdf"))
#
# word_app = None
#
# try:
#     # Run the extracted engine function
#     word_app = process_change_orders(
#         excel_file=EXCEL_FILE,
#         template_doc=TEMPLATE_DOC,
#         output_dir=OUTPUT_DIR,
#         master_docx_path=MASTER_DOCX_PATH,
#         master_pdf_path=MASTER_PDF_PATH,
#         test_sample_pdf_path=TEST_SAMPLE_PDF_PATH
#     )
#
#     print("\n✅ Success! Each letter inside the master document has been stamped with its own unique barcode.")
#
# except FileNotFoundError:
#     print(f"❌ Error: Missing assets. Ensure paths to Excel and Word are correct.")
# except Exception as e:
#     print(f"❌ A fatal runtime disruption occurred: {e}")
#
# finally:
#     # Clean up Word COM automation state safely
#     if word_app is not None:
#         try:
#             word_app.Quit()
#             print("Word Application shutdown complete.")
#         except Exception:
#             pass
