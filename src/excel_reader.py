import pandas as pd
from pathlib import Path

REQUIRED_COLUMNS = [
    "Model_ID",
    "Mfg_Packer",
    "Vertical",
    "Support_Number",
    "FSN",
    "MRP",
    "Net_Quantity",
    "Email",
    "Color",
    "Origin",
    "Mfg_M/Y",
    "Brand",
    "Qty",
    "EAN_No",
]


def read_excel(path, sheet_name):
    file_path = Path(path)

    if not file_path.exists():
        raise FileNotFoundError(f"Excel file not found: {file_path}")

    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df.columns = df.columns.str.strip()

    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    records = []

    for idx, row in df.iterrows():
        record = {}

        try:
            # ----------------------------
            # Basic Field Extraction
            # ----------------------------
            for col in REQUIRED_COLUMNS:
                value = str(row[col]).strip()
                if value == "" or value.lower() == "nan":
                    raise ValueError(f"Empty value in column '{col}'")
                record[col] = value

            # ----------------------------
            # MRP Formatting
            # Rs.999 (Inclusive of all taxes)
            # ----------------------------
            mrp_raw = record["MRP"]
            mrp_clean = (
                mrp_raw
                .replace("Rs.", "")
                .replace("₹", "")
                .replace("/-", "")
                .strip()
            )
            mrp_value = int(float(mrp_clean))
            record["MRP"] = f"Rs.{mrp_value} (Inclusive of all taxes)"

            # ----------------------------
            # Date Formatting (Mfg_M/Y)
            # Convert to: Dec 2025
            # ----------------------------
            

            # ----------------------------
            # Qty Validation
            # ----------------------------
            qty_value = int(float(record["Qty"]))
            record["Qty"] = qty_value

            if qty_value != 8:
                print(f"[WARNING] Row {idx + 2}: Qty is {qty_value} (expected 8)")

            # ----------------------------
            # Clean EAN
            # ----------------------------
            record["EAN_No"] = record["EAN_No"].replace(" ", "")

            records.append(record)

        except Exception as e:
            print(f"[SKIPPED ROW {idx + 2}] {e}")

    return records
