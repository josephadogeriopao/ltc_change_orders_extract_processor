import os
import sys
from dotenv import load_dotenv

# Import the centralized constants and pipeline runner
from models.constants import PropertyType
from pipeline import run_pipeline

def main():
    # 1. Automatically load configurations into environment workspace memory
    load_dotenv()

    # 2. Extract configuration tokens with programmatic fallback protections
    data_file_raw = os.getenv("DATA_FILE")
    output_dir_raw = os.getenv("OUTPUT_DIR")
    property_type_raw = os.getenv("PROPERTY_TYPE")  # e.g., "PERSONAL" or "REAL"

    # Safe structural assertion step ensuring no key metrics are missing
    if not all([data_file_raw, output_dir_raw, property_type_raw]):
        print("❌ Critical System Launch Fault: Active variables mapping is incomplete in .env configuration layout.")
        print("👉 Ensure DATA_FILE, OUTPUT_DIR, and PROPERTY_TYPE are defined.")
        sys.exit(1)

    # 3. Resolve absolute paths safely
    target_data_file = os.path.abspath(data_file_raw)
    chosen_base_dir = os.path.abspath(output_dir_raw)

    # 4. Map the string from .env to the correct PropertyType Enum class reference
    try:
        target_property_type = PropertyType[property_type_raw.upper()]
    except KeyError:
        valid_types = [t.name for t in PropertyType]
        print(f"❌ Launch Fault: Invalid PROPERTY_TYPE '{property_type_raw}' provided in .env file.")
        print(f"👉 Choose one of the following exact options: {valid_types}")
        sys.exit(1)

    # 5. File target execution guard verification step
    if not os.path.exists(target_data_file):
        print(f"❌ Launch Fault: Source data file does not exist at location: {target_data_file}")
        sys.exit(1)

    # 6. Boot the detached pipeline runner context loop
    run_pipeline(
        excel_file_path=target_data_file,
        property_type=target_property_type,
        base_output_dir=chosen_base_dir
    )

if __name__ == "__main__":
    main()
