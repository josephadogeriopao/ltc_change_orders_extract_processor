import os
import sys
import logging
from models.constants import PropertyType, JOB_REGISTRY
from utils.mailmerge import process_change_orders

# Access the centralized live console logger instance
logger = logging.getLogger("AppEngine")


def generate_output_paths(excel_file_path: str, property_type: PropertyType, base_output_dir: str) -> tuple[str, str, str, str]:
    """
    Utility function that extracts the source spreadsheet filename and generates
    the distinct absolute paths required by the change order engine.
    """
    # 1. Fetch matching job blueprints from configuration map
    config = JOB_REGISTRY[property_type]

    # 2. Extract base file name from the source spreadsheet path
    excel_file = os.path.abspath(excel_file_path)
    file_basename, _ = os.path.splitext(os.path.basename(excel_file))

    # 3. Dynamic Subdirectory Construction
    output_dir = os.path.normpath(os.path.join(os.path.abspath(base_output_dir), config.folder_suffix))

    # --- DIRECTORY SANITIZATION FIX ---
    # Create the output target folders immediately if they do not exist
    os.makedirs(output_dir, exist_ok=True)
    # ----------------------------------

    # 4. Generate strict filename formatting parameters
    pfx = config.file_prefix
    disp = config.display_name

    master_docx_name = f"{pfx} Mailmerge_{file_basename}.docx"
    test_sample_name = f"test_{disp}_{file_basename}.pdf"
    master_pdf_name = f"{pfx}_{file_basename}.pdf"

    # 5. Resolve fully absolute paths
    master_docx_path = os.path.normpath(os.path.join(output_dir, master_docx_name))
    master_pdf_path = os.path.normpath(os.path.join(output_dir, master_pdf_name))
    test_sample_pdf_path = os.path.normpath(os.path.join(output_dir, test_sample_name))

    return output_dir, master_docx_path, master_pdf_path, test_sample_pdf_path


def run_pipeline(excel_file_path: str, property_type: PropertyType, base_output_dir: str, ui_app=None) -> None:
    """Orchestrates configuration path generation and feeds it into the mailmerge runner loop."""

    # 1. Use the utility function to parse the filename and generate absolute target paths
    output_dir, master_docx_path, master_pdf_path, test_sample_pdf_path = generate_output_paths(
        excel_file_path=excel_file_path,
        property_type=property_type,
        base_output_dir=base_output_dir
    )

    # Convert print statements to use the dynamic live UI console logger streams
    logger.info(f"Initializing Framework Run | Target Type: [ {property_type.name} ]")
    logger.info(f"Source Tracker Sheet Path: {os.path.basename(excel_file_path)}")
    logger.info(f"Target Output Folder Location: {output_dir}")

    word_app = None

    try:
        # Resolve absolute template path relative to where constants.py is safely stored
        raw_template_path = JOB_REGISTRY[property_type].template_path

        # 2. Trigger framework processing engine exactly once
        word_app = process_change_orders(
            excel_file=os.path.abspath(excel_file_path),
            template_doc=os.path.abspath(raw_template_path),
            output_dir=output_dir,
            master_docx_path=master_docx_path,
            master_pdf_path=master_pdf_path,
            test_sample_pdf_path=test_sample_pdf_path,
            ui_app=ui_app
        ) # 👈 FIXED: Closed the parenthesis safely here!

        logger.info("Success! Pipeline completed layout mapping tasks successfully.")

    except FileNotFoundError as fnf:
        logger.error(f"Processing dropped due to missing data file assets: {fnf}")
        raise fnf
    except Exception as e:
        logger.error(f"A fatal runtime disruption occurred during processing loop: {e}")
        raise e

    finally:
        # 3. Clean up Word COM automation state safely
        if word_app is not None:
            try:
                word_app.Quit()
                logger.info("Word Application background wrapper shutdown complete.")
            except Exception:
                pass
