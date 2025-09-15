import pandas as pd
from Maid_Trinket import DeepClean
from Builder_Trinket import DeetsBuilder

# Define your header aliases (input → normalized name)
HEADER_ALIASES = {
    "ProductID|SKU": "SKU",
    "Quantity|QTY": "QTY",
    "Duration|INITIAL DURATION": "INITIAL DURATION",
    "PricingTerm|PRICED PER X": "PRICED PER X",
    "UnitListPrice|UNIT LIST PRICE": "UNIT LIST PRICE",
    "ExtendedListPrice|EXTENDED LIST PRICE": "EXTENDED LIST PRICE",
    "Discount|DISCOUNT % OFF LIST": "DISCOUNT % OFF LIST",
    "UnitCost|UNIT COST": "UNIT COST",
    "ExtendedNetCost|EXTENDED NET COST": "EXTENDED NET COST"
}

# Final headers you want in the output
FINAL_HEADERS = [
    "SKU", "QTY", "INITIAL DURATION", "PRICED PER X",
    "UNIT LIST PRICE", "EXTENDED LIST PRICE",
    "DISCOUNT % OFF LIST", "UNIT NET PRICE", "EXTENDED NET PRICE (months)"
]

def main():
    InputPath = "input.xlsx"       # your test input
    CleanedPath = "output_clean.xlsx"
    FinalPath = "output_final.xlsx"

    print("✨ Step 1: Running Cleaner_Trinket...")
    DeepClean(InputPath, CleanedPath)
    print(f"✅ Cleaning complete! Saved as {CleanedPath}")

    print("✨ Step 2: Loading cleaned Details tab...")
    df = pd.read_excel(CleanedPath, sheet_name="Details")

    # Example user input
    UserInput = {
        "PricingType": "HOLD BACK",   # or MARKUP, MARGIN
        "PercentInput": 2          # e.g. 10% holdback
    }

    print("✨ Step 3: Building Details with maths...")
    BobTheBuilder = DeetsBuilder(df, UserInput, HEADER_ALIASES, FINAL_HEADERS)
    BobTheBuilder.MakeHeadersGreatAgain()
    BobTheBuilder.DoMaths()
    BobTheBuilder.YouDontEvenGoHere()
    FinalDF = BobTheBuilder.Finalize()

    print("✨ Step 4: Saving final output...")
    with pd.ExcelWriter(FinalPath, engine="openpyxl") as writer:
        FinalDF.to_excel(writer, sheet_name="Details", index=False)

    print(f"✅ All done! Final file saved as {FinalPath}")

if __name__ == "__main__":
    main()
