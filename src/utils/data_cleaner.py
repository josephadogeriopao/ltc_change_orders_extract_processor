import pandas as pd

NUMERIC_ID_FIELDS = {'TAX_YEAR', 'BATCH_NO', 'BATCH_ITEM_NO'}
DATE_FIELDS = ['STATUS_DATE', 'BATCH_SUBMITTED']


def clean_value(key_str: str, val) -> str:
    """Applies specific data type formatting rules to values."""
    if pd.isna(val) or val == "":
        return ""

    if isinstance(val, (int, float)):
        if key_str in NUMERIC_ID_FIELDS:
            return str(int(val))
        if val % 1 == 0:
            return f"{int(val):,}"

        formatted_val = f"{val:,.2f}"
        return formatted_val[:-3] if formatted_val.endswith(".00") else formatted_val

    return str(val).strip()


def pipeline_clean_records(raw_records: list[dict]) -> tuple[list[dict], list[str]]:
    """Cleans names, normalizes addresses, and extracts parcel list references."""
    cleaned_records = []
    parcel_ids = []

    for record in raw_records:
        clean_record = {str(k).strip(): clean_value(str(k).strip(), v) for k, v in record.items()}

        # Mirror case-variants for MailMerge placeholders
        reason = clean_record.get("REASON", "")
        clean_record.update({"reason": reason, "Reason": reason})

        parcel_id = clean_record.get("PARCELID", "").strip()
        clean_record.update({"parcelid": parcel_id, "PARCELID": parcel_id})
        parcel_ids.append(parcel_id)

        # Fix structural layout spacing anomalies for addresses
        if clean_record.get("TAXPAYER_ADDR2") == "":
            clean_record["TAXPAYER_ADDR2"] = clean_record.get("TAXPAYER_ADDR3", "")
            clean_record["TAXPAYER_ADDR3"] = ""

        # Normalize clean date format splits
        for date_col in DATE_FIELDS:
            if clean_record.get(date_col) and " " in clean_record[date_col]:
                clean_record[date_col] = clean_record[date_col].split(" ")[0]

        cleaned_records.append(clean_record)

    return cleaned_records, parcel_ids
