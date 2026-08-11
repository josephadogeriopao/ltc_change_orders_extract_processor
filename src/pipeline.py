import os
from models.constants import PropertyType, JOB_REGISTRY
from utils.mailmerge import process_change_orders


def generate_output_paths(excel_file_path: str, property_type: PropertyType, base_output_dir: str) -> tuple[
    str, str, str, str]:
    """
    Utility function that extracts the source spreadsheet filename and generates
    the distinct absolute paths required by the change order engine.
    """
    # 1. Fetch matching job blueprints from configuration map
    config = JOB_REGISTRY[property_type]

    # 2. Extract base file name from the source spreadsheet path (drops .xls or .xlsx)
    excel_file = os.path.abspath(excel_file_path)
    file_basename, _ = os.path.splitext(os.path.basename(excel_file))

    # 3. Dynamic Subdirectory Construction
    output_dir = os.path.normpath(os.path.join(os.path.abspath(base_output_dir), config.folder_suffix))

    # 4. Generate strict filename formatting parameters
    pfx = config.file_prefix  # e.g., "PP" or "REAL"
    disp = config.display_name  # e.g., "Personal_Property" or "Real_Property"

    master_docx_name = f"{pfx} Mailmerge_{file_basename}.docx"
    test_sample_name = f"test_{disp}_{file_basename}.pdf"
    master_pdf_name = f"{pfx}_{file_basename}.pdf"

    # 5. Resolve fully absolute paths
    master_docx_path = os.path.normpath(os.path.join(output_dir, master_docx_name))
    master_pdf_path = os.path.normpath(os.path.join(output_dir, master_pdf_name))
    test_sample_pdf_path = os.path.normpath(os.path.join(output_dir, test_sample_name))

    return output_dir, master_docx_path, master_pdf_path, test_sample_pdf_path


def run_pipeline(excel_file_path: str, property_type: PropertyType, base_output_dir: str) -> None:
    """Orchestrates configuration path generation and feeds it into the mailmerge runner loop."""

    # 1. Use the utility function to parse the filename and generate absolute target paths
    output_dir, master_docx_path, master_pdf_path, test_sample_pdf_path = generate_output_paths(
        excel_file_path=excel_file_path,
        property_type=property_type,
        base_output_dir=base_output_dir
    )

    print("=" * 80)
    print(f"🚀 Initializing Framework Run | Property Type Target: [ {property_type.name} ]")
    print(f"   > Data Source Target:  {os.path.abspath(excel_file_path)}")
    print(f"   > Target Output Dir:   {output_dir}")
    print(f"   > Target Docx Output:  {os.path.basename(master_docx_path)}")
    print(f"   > Target Sample PDF:   {os.path.basename(test_sample_pdf_path)}")
    print(f"   > Target Master PDF:   {os.path.basename(master_pdf_path)}")
    print("=" * 80)

    word_app = None

    try:
        # 2. Trigger framework processing engine exactly once
        word_app = process_change_orders(
            excel_file=os.path.abspath(excel_file_path),
            template_doc=JOB_REGISTRY[property_type].template_path,
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
        # 3. Clean up Word COM automation state safely
        if word_app is not None:
            try:
                word_app.Quit()
                print("🔒 Word Application background wrapper shutdown complete.")
            except Exception:
                pass
