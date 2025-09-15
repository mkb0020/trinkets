import pandas as pd
import tkinter as tk
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.drawing.image import Image
from Maths_Trinket import MathsPreReqs, Maths
from Maid_Trinket import Floaties
from Styles_Trinket import Decimals



class DeetsBuilder:
    def __init__(self, df: pd.DataFrame, UserInput: dict, HeaderAliases: dict, FinalHeaders: list):
        """
        df: cleaned dataframe from Maid_Trinket
        user_input: dict with PricingType + PercentInput
        header_aliases: mapping of possible header variations → standardized header
        final_headers: list of final column order for output
        """
        self.df = df.copy()
        self.UserInput = UserInput
        self.HeaderAliases = HeaderAliases
        self.FinalHeaders = FinalHeaders
        self.Maths = Maths()
        self.MathsPreReqs = MathsPreReqs()  # we only use its math helpers
        



    def MakeHeadersGreatAgain(self):
        RrenameMap = {}
        for aliases, standard_name in self.HeaderAliases.items():
            for alias in aliases.split("|"):
                if alias in self.df.columns:
                    RrenameMap[alias] = standard_name
        self.df = self.df.rename(columns=RrenameMap)

    def DoMaths(self): #Apply pricing row by row
        PricingType = self.UserInput.get("PricingType", "").upper()
        PercentFraction = self.MathsPreReqs.GetPercent(self.UserInput.get("PercentInput", 0))

        Percent = self.UserInput["PercentInput"] / 100
        def RowMaths(row):
            QTY = round(int(row.get("QTY", 0)), 0)
            LineDuration = round(row.get("INITIAL DURATION", 2), 2)
            PTerm = round(row.get("PRICED PER X", 0), 0)
            UnitList = Floaties(row.get("UNIT LIST PRICE", 2), decimals=2)
            UnitCost = Floaties(row.get("UNIT COST", 2), decimals=2)
            LineDiscount = Floaties(row.get("DISCOUNT % OFF LIST", 2), decimals=2)
            Discount = Floaties(self.MathsPreReqs.GetDiscount(PricingType, PercentFraction, LineDiscount, UnitCost, UnitList), decimals=2)
            print(f"🧮 Final Discount Used for Pricing: {Discount:.2%}")
            UnitNP = Floaties(self.Maths.GetUnitNP(Discount, UnitList), decimals=2)
            print(f"💰 Unit Net Price: {UnitNP}")
            LineExtendedNP = Floaties(self.Maths.GetLineExtendedNP(UnitNP, QTY, LineDuration, PTerm), decimals=2)


            return pd.Series({
                "UNIT NET PRICE": UnitNP,
                "EXTENDED NET PRICE (months)": LineExtendedNP,
                "DISCOUNT % OFF LIST": Discount,
            })

        maths_df = self.df.apply(RowMaths, axis=1)
        self.df = pd.concat(
            [self.df.drop(columns=maths_df.columns.intersection(self.df.columns)), maths_df],
            axis=1
        )
        print("✅ Maths applied. Final columns:", self.df.columns.tolist())
    


    def YouDontEvenGoHere(self): # drop unwanted columns
        ByeBye = [
            "UNIT COST", "EXTENDED NET COST"
        ]
        self.df.columns = self.df.columns.str.strip()
        self.df = self.df.drop(columns=[col for col in ByeBye if col in self.df.columns])

    def Finalize(self):
        """Reorder columns and keep only what we care about."""
        self.df = self.df[[col for col in self.FinalHeaders if col in self.df.columns]]
        self.df = Decimals(self.df, ["UNIT NET PRICE", "EXTENDED NET PRICE (months)", "DISCOUNT % OFF LIST"])
        return self.df