import os
import sys
from dotenv import load_dotenv

# Import the orchestrator function from your modular structure
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
    print("🚀 Initializing Change Order Framework Run Instance")
    print(f"   > Data Source Target:  {data_file}")
    print(f"   > Output Destination: {output_dir}")
    print("=" * 80)

    word_app = None

    try:
        # 4. Trigger processing framework engine exactly once
        # This will return the word_app instance for managed cleanup below
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
