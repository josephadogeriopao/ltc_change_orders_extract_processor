import os
import sys

from dotenv import load_dotenv

from src.models.owner import Owner
from src.utils.convert import convert_txt_to_excel

# Load the variables from the local .env file
load_dotenv()

txt_file = os.environ.get("TXT_FILE", r"PATH TO PDF FILE")

def main() -> None:
    # Defensive check: Ensure both environment variables exist before running
    if not all([txt_file]):
        print("❌ Error: Missing configuration paths in your .env file or fallbacks!")
        return

    print("\nStarting Generating LTC Change Orders Service...")
    print(f"-> text file: {txt_file}")
    print("Finished Generating LTC Change Orders Service...")

    saved_at : str = convert_txt_to_excel(txt_file)

    print(f"File automatically saved to: {saved_at}")


if __name__ == "__main__":
    main()
    sys.exit(0)
