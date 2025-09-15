import pandas as pd

DEETS_HEADER_MAP = { # Mapping of original headers → normalized headers
    "ProductID": "PRODUCT",
    "SKU": "SKU",
    "Quantity": "QTY",
    "Duration": "INITIAL DURATION",
    "PricingTerm": "PRICED PER X",
    "UnitListPrice": "UNIT LIST PRICE",
    "ExtendedListPrice": "EXTENDED LIST PRICE",
    "Discount": "DISCOUNT % OFF LIST",
    "UnitCost": "UNIT COST",
    "ExtendedNetCost": "EXTENDED NET COST"
}

def CleanDeetsTab(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clean the details tab dataframe:
    - Drop fully blank rows
    - Drop rows with '--' in important columns
    - Strip and rename headers to normalized names
    """
    df = df.dropna(how="all") # Drop fully blank rows

    df.columns = [str(c).strip() for c in df.columns] # Strip whitespace from column names

    CheckColumns = ["SKU", "UnitListPrice", "Discount"]  # Drop rows where SKU, UnitListPrice, or Discount contain '--'
    for col in CheckColumns:
        if col in df.columns:
            df = df[df[col] != "--"]

    df = df.rename(columns={col: DEETS_HEADER_MAP.get(col, col) for col in df.columns}) # Rename headers

    return df.reset_index(drop=True)


def DeepClean(InputPath: str, OutputPath: str):
    """
    Clean an input Excel file and save to a new file.
    - Keeps Summary tab as-is
    - Cleans the Details tab
    """
    # Read the workbook
    xls = pd.ExcelFile(InputPath)

    # First sheet = Summary (keep as-is, no header)
    SummaryDF = pd.read_excel(xls, sheet_name="Summary", header=None)

    # Second sheet = Details (reference ID name, skip first useless row)
    DeetsTab = xls.sheet_names[1]
    DeetsDF = pd.read_excel(xls, sheet_name=DeetsTab, header=1)

    # Clean details tab
    DeetsDF = CleanDeetsTab(DeetsDF)

    # Write cleaned data back to Excel
    with pd.ExcelWriter(OutputPath, engine="openpyxl") as writer:
        SummaryDF.to_excel(writer, sheet_name="Summary", index=False, header=False)
        DeetsDF.to_excel(writer, sheet_name="Details", index=False)


def Floaties(val, decimals=None): #safe float conversion
    try:
        result = float(val)
        return round(result, decimals) if decimals is not None else result
    except (TypeError, ValueError):
        return 0.0
