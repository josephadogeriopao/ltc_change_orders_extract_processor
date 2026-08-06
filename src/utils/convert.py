# src/converter.py
import csv
import os
import openpyxl


class TildeDialect(csv.excel):
    delimiter = ','
    quotechar = '~'
    quoting = csv.QUOTE_MINIMAL


csv.register_dialect('tilde_csv', TildeDialect)


# Notice: we removed 'excel_file_path' as an input parameter
def convert_txt_to_excel(txt_file_path):
    """Parses text rows into Excel, dynamically generating an identical file name with a .xlsx extension."""

    # 1. Validation
    if not os.path.exists(txt_file_path):
        raise FileNotFoundError(f"The input file could not be found: {txt_file_path}")

    # 2. Dynamically build the matching Excel file path
    # os.path.splitext("/path/to/file.txt") -> ("/path/to/file", ".txt")
    base_path, _ = os.path.splitext(txt_file_path)
    excel_file_path = base_path + ".xlsx"

    # Ensure the target folder exists (if the txt file was in a subfolder)
    output_dir = os.path.dirname(excel_file_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    # 3. Map field names to their strict Oracle VARCHAR2 string limits
    FIELD_LIMITS = {
        "reject_reason": 255,
        "ltc_comment": 100,
        "reason": 255,
        "prop_desc": 255
    }

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tax Data"
    ws.sheet_view.showGridLines = True

    with open(txt_file_path, mode='r', encoding='utf-8', newline='') as f:
        reader = csv.reader(f, dialect='tilde_csv')

        exclude_index = None
        header_mapping = {}  # Tracks column index positions to field names

        for row_index, row in enumerate(reader, start=1):
            # Enforce clean, predictable row height (No text wrapping overflow)
            ws.row_dimensions[row_index].height = 20

            # 4. Analyze the header row
            if row_index == 1:
                clean_headers = [h.strip().lower() for h in row]

                # Identify 'property_id' position to exclude it
                if "property_id" in clean_headers:
                    exclude_index = clean_headers.index("property_id")

                # Build an index map to recognize fields in subsequent data rows
                header_mapping = {i: name for i, name in enumerate(clean_headers)}

            # Track independent target columns in Excel to avoid blank gaps
            excel_col_counter = 1

            # 5. Process and write cells
            for col_index, value in enumerate(row):
                # Skip 'property_id' entirely
                if exclude_index is not None and col_index == exclude_index:
                    continue

                clean_value = value.strip() if value else ""

                # Only apply character truncation to data rows (skip the headers themselves)
                if row_index > 1:
                    field_name = header_mapping.get(col_index)
                    if field_name in FIELD_LIMITS:
                        max_len = FIELD_LIMITS[field_name]
                        # Safe slice: extracts text up to the max limit without crashing if short
                        clean_value = clean_value[:max_len]

                ws.cell(row=row_index, column=excel_col_counter, value=clean_value)
                excel_col_counter += 1

    wb.save(excel_file_path)
    return excel_file_path  # Returns the new path back so your app knows where it saved
