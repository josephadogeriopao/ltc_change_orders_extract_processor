import json
import pandas as pd


def load_raw_records(file_path: str) -> list[dict]:
    """Loads records dynamically from JSON, XLS, or XLSX files."""
    file_path_lower = file_path.lower()

    if file_path_lower.endswith(".json"):
        with open(file_path, "r", encoding="utf-8") as f:
            raw_records = json.load(f)
        if isinstance(raw_records, dict):
            return [raw_records]
        return raw_records

    # Handle Excel variants
    chosen_engine = "openpyxl" if file_path_lower.endswith(".xlsx") else "xlrd"
    df = pd.read_excel(file_path, engine=chosen_engine, header=0)
    df.columns = df.columns.str.strip()
    return df.to_dict(orient="records")
