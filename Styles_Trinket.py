import pandas as pd
import tkinter as tk
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.drawing.image import Image
from openpyxl import load_workbook

class DrippyKit:
    # ---------------------------- NUMBER FORMATS ---------------------------- 
    StandardCurrencyCols = {"UNIT LIST PRICE", "UNIT NET PRICE", "EXTENDED NET PRICE"}
    TDSummCurrencyCols = {"UNIT LIST PRICE", "UNIT NET PRICE", "ESTIMATED CREDIT", "ESTIMATED INVOICE",	"TRUE DELTA NET COST"}
    TDDeetsCurrencyCols = {"UNIT LIST PRICE", "UNIT NET PRICE","EXISTNG NET PRICE", "NEW NET PRICE", "EXISTING PRORATED PRICE", "NEW PRORATED PRICE", "ESTIMATED CREDIT", "ESTIMATED INVOICE", "TRUE DELTA NET COST"}
    NewSummaryCurrencyCols = StandardCurrencyCols #placehlder
    NewDeetsCurrencyCols = StandardCurrencyCols #placehlder
    RenewalSummaryCurrencyCols = StandardCurrencyCols #placehlder
    RenewalDeetsCurrencyCols = StandardCurrencyCols #placehlder
    ModSummaryCurrencyCols = StandardCurrencyCols #placehlder
    ModDeetsCurrencyCols = StandardCurrencyCols #placehlder 
    # ---------------------------- FILLS ---------------------------- 
    HeaderColor = PatternFill(start_color="4B1395", end_color="4B1395", fill_type="solid")
    SubHeaderColor = PatternFill(start_color="CBCEDF", end_color="CBCEDF", fill_type="solid")
    BodyColor = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    # ---------------------------- BORDERS ---------------------------- 
    Border_IttyBitty = Side(style="thin", color="000000")
    Border_SheThicc = Side(style="thick", color="000000")
    Border_DoubleBottom = Side(border_style="double", color="000000")
    # ---------------------------- FONTS ----------------------------        
    MainHeaderFont = Font(name="Aptos Display", bold=True, color="FFFFFF", size=14)
    HeaderFont = Font(name="Aptos Narrow", bold=True, color="FFFFFF", size=10)
    SubHeaderFont = Font(name="Aptos Narrow", bold=True, color="000000", size=10)
    BodyFont = Font(name="Aptos Narrow", bold=False, color="000000", size=10)
    BodyBoldFont = Font(name="Aptos Narrow", bold=True, color="000000", size=10)
    # ---------------------------- ALIGNMENTS ---------------------------- 
    Middle = Alignment(horizontal="center", vertical="center", wrap_text=True)
    Lefty = Alignment(horizontal="left", vertical="center", indent=1)
    Righty = Alignment(horizontal="right", vertical="center", indent=1)

    @staticmethod
    def ApplyDrip(cell, font=None, alignment=None, fill=None, border=None):
        if font:
            cell.font = font
        if alignment:
            cell.alignment = alignment
        if fill:
            cell.fill = fill
        if border:
            cell.border = border

    @staticmethod
    def apply_thicc_outline(ws, FirstRow, LastRow, FirstColumn, LastColumn, ThiccBorder=Border_SheThicc):
        for row in ws.iter_rows(min_row=FirstRow, max_row=LastRow, min_col=FirstColumn, max_col=LastColumn):
            for cell in row:
                borders = {k: cell.border.__dict__[k] for k in ("left", "right", "top", "bottom")}
                if cell.row == 1: borders["top"] = ThiccBorder
                if cell.row == LastRow: borders["bottom"] = ThiccBorder
                if cell.column == 1: borders["left"] = ThiccBorder
                if cell.column == LastColumn: borders["right"] = ThiccBorder
                cell.border = Border(**borders)


class DeetsDrip:
    def __init__(self, filepath, sheet_name):
        self.filepath = filepath
        self.sheet_name = sheet_name
        self.wb = load_workbook(filepath)
        self.ws = self.wb[sheet_name]
        LastColumn = self.ws.max_column
        HeaderRow = next(self.ws.iter_rows(min_row=1, max_row=1))
        FirstItemsRow = 2  # Just use the row number
        LastRow = self.ws.max_row + 1
        self.ws.row_dimensions[1].height = 45
        for cell in HeaderRow:
            DrippyKit.ApplyDrip(
            cell,
            font=DrippyKit.HeaderFont,
            alignment=DrippyKit.Middle,
            fill=DrippyKit.HeaderColor,
            border=DrippyKit.Border_IttyBitty
        )   
        
        for row in range(FirstItemsRow, LastRow):
            self.ws.row_dimensions[row].height = 15
            for col in range(1, LastColumn):
                cell = self.ws.cell(row, col)
                DrippyKit.ApplyDrip(
                cell,
                font=DrippyKit.BodyFont,
                alignment=DrippyKit.Middle,
                fill=DrippyKit.BodyColor,
                border=DrippyKit.Border_IttyBitty
            )   
            
        DrippyKit.apply_thicc_outline(self.ws, 1, LastRow, 1, LastColumn)

        self.wb.save(self.filepath)


class SumaryDrip:
    def __init__(self, filepath, sheet_name):
        self.filepath = filepath
        self.sheet_name = sheet_name
        self.wb = load_workbook(filepath)
        self.ws = self.wb[sheet_name]

        LastColumn = self.ws.max_column
        MainHeaderRow = next(self.ws.iter_rows(min_row=1, max_row=1))
        FirstGenInfoRow = 3
        LastGenInfoRow = 19 #this will need to be dynamic but we'll worry about that later!
        HeaderRow = LastGenInfoRow + 2
        FirstItemsRow = HeaderRow + 1
        LastItemsRow = self.ws.max_row - 5
        TotalsRow = LastItemsRow + 1
        FirstNotesRow = TotalsRow + 2
        LastNotesRow = FirstNotesRow + 3
       
        for cell in MainHeaderRow: #Style title header
            DrippyKit.ApplyDrip(
            cell,
            font=DrippyKit.MainHeaderFont,
            alignment=DrippyKit.Middle,
            fill=DrippyKit.HeaderColor,
            border=DrippyKit.Border_IttyBitty
        )   
        self.ws.row_dimensions[1].height = 22
        DrippyKit.apply_thicc_outline(self.ws, 1, 1, 1, LastColumn) #Apply Thick utline around main title header
        
        self.ws.row_dimensions[2].height = 5 #Row between the main title header and the general info section
        #-------------GENERAL INFO SECTION-------------
        for row in range(FirstGenInfoRow, LastGenInfoRow): #Style General Info Section
            self.ws.row_dimensions[row].height = 12
            for col in range(1): #Column A in the General Info Section
                cell = self.ws.cell(row, col)
                DrippyKit.ApplyDrip(
                    cell,
                    font=DrippyKit.SubHeaderFont,
                    alignment=DrippyKit.Lefty,
                    fill=DrippyKit.SubHeaderColor,
                    border=DrippyKit.Border_IttyBitty
                )
            for col in range(2): #Column B in the General Info Section
                    cell = self.ws.cell(row, col)
                    DrippyKit.ApplyDrip(
                    cell,
                    font=DrippyKit.BodyFont,
                    alignment=DrippyKit.Righty,
                    fill=DrippyKit.BodyColor,
                    border=DrippyKit.Border_IttyBitty
                )
        DrippyKit.apply_thicc_outline(self.ws, FirstGenInfoRow, LastGenInfoRow, 1, 2) #Apply Thick Outline to General Info Section

        self.ws.row_dimensions[HeaderRow - 1].height = 5 #Row between the General Info and the items section
        #-------------SUMARY ITEMS SECTION-------------
        for cell in HeaderRow: #Style Items Header Section
            DrippyKit.ApplyDrip(
            cell,
            font=DrippyKit.HeaderFont,
            alignment=DrippyKit.Middle,
            fill=DrippyKit.HeaderColor,
            border=DrippyKit.Border_IttyBitty
        )   
        self.ws.row_dimensions[HeaderRow].height = 30
                    
        for row in range(FirstItemsRow, LastItemsRow): #Style Items Section
            self.ws.row_dimensions[row].height = 15
            for col in range(1, LastColumn):
                cell = self.ws.cell(row, col)
                DrippyKit.ApplyDrip(
                cell,
                font=DrippyKit.BodyFont,
                alignment=DrippyKit.Middle,
                fill=DrippyKit.BodyColor,
                border=DrippyKit.Border_IttyBitty
            )   

        for cell in TotalsRow: #Style Totals Row
            DrippyKit.ApplyDrip(
            cell,
            font=DrippyKit.SubHeaderFont,
            alignment=DrippyKit.Lefty,
            fill=DrippyKit.SubHeaderColor,
            border=DrippyKit.Border_IttyBitty
            )
        self.ws.row_dimensions[TotalsRow].height = 15

        DrippyKit.apply_thicc_outline(self.ws, HeaderRow, TotalsRow, 1, LastColumn) #Thick outline around items section - including the headers and totals row

        self.ws.row_dimensions[TotalsRow + 1].height = 5 #Row between the totals row and notes section
        #-------------NOTES SECTION-------------
        for cell in FirstNotesRow: #Style Notes Header
                DrippyKit.ApplyDrip(
                cell,
                font=DrippyKit.HeaderFont,
                alignment=DrippyKit.Lefty,
                fill=DrippyKit.HeaderColor,
                border=DrippyKit.Border_IttyBitty
            )
        self.ws.row_dimensions[FirstNotesRow].height = 15
        for row in range(FirstNotesRow+1, LastNotesRow): #Style Notes Section
            self.ws.row_dimensions[row].height = 15
            for col in range(1, LastColumn):
                cell = self.ws.cell(row, col)
                DrippyKit.ApplyDrip(
                cell,
                font=DrippyKit.BodyFont,
                alignment=DrippyKit.Lefty,
                fill=DrippyKit.SubHeaderColor,
                border=DrippyKit.Border_IttyBitty
            )
                
        DrippyKit.apply_thicc_outline(self.ws, FirstNotesRow, LastNotesRow, 1, LastColumn) #Thick outline around notes section

        DrippyKit.apply_thicc_outline(self.ws, 1, LastNotesRow, 1, LastColumn) #Thick outline around The whole thing

        self.wb.save(self.filepath)


    





    def summary_main_header_swag(self):
        cell = self.ws["A1"]
        cell.alignment = Middle
        cell.fill = HeaderColor
        cell.font = MainHeaderFont
        cell.border = Border(left=Border_SheThicc, right=Border_SheThicc, top=Border_SheThicc, bottom=Border_DoubleBottom)
        self.ws.row_dimensions[1].height = 22
        self.ws.row_dimensions[2].height = 5

class DeltaDeetsDrip:
    def __init__(self, filepath, sheet_name):
        self.filepath = filepath
        self.sheet_name = sheet_name
        self.wb = load_workbook(filepath)
        self.ws = self.wb[sheet_name]

# Alignment rules
        self.Leftys = {"SKU", "PRODUCT DESCRIPTION"}
        self.Middles = {
            "PACKAGE", "TIER", "LICENSE SHIFT ELIGIBLE", "ALLOCATION STATUS",
            "EXISTING QTY", "NEW QTY", "(+)LICENSE SHIFT", "(-)LICENSE SHIFT", "DELTA QTY"
        }
        self.Rightys = {
            "UNIT LIST PRICE", "DISCOUNT OFF LIST", "UNIT NET PRICE",
            "EXISTNG NET PRICE (months)", "NEW NET PRICE (months)",
            "EXISTING PRORATED PRICE", "NEW PRORATED PRICE",
            "ESTIMATED CREDIT", "ESTIMATED INVOICE", "TRUE DELTA NET COST"
        }

        self.CurrencyColumns = TDSummCurrencyCols
        self.PercentColumns = ["DISCOUNT OFF LIST"]

    def delta_deets_column_widths(self):
        for col in ["A", "E"]:
            self.ws.column_dimensions[col].width = 13
        for col in ["B", "C", "F"]:
            self.ws.column_dimensions[col].width = 20
        for col in ["D", "G", "H", "I", "K", "L", "M", "N"]:
            self.ws.column_dimensions[col].width = 12
        for col in ["O", "P", "Q", "R", "S", "T", "U"]:
            self.ws.column_dimensions[col].width = 15

    def delta_deets_number_formats(self):
        header = [cell.value for cell in self.ws[1]]
        for col_name in self.CurrencyColumns:
            if col_name in header:
                col_idx = header.index(col_name) + 1
                for row in self.ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE

        for col_name in self.PercentColumns:
            if col_name in header:
                col_idx = header.index(col_name) + 1
                for row in self.ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        cell.number_format = "0.00%"

    def delta_deets_dynamic_headers(self):
        # Load durations from Summary tab
        summary_ws = self.wb["Summary"]
        agreement_duration = summary_ws["B9"].value or "?"
        remaining_duration = summary_ws["B10"].value or "?"

        self.ws["O1"].value = f"EXISTING\nNET PRICE\n({agreement_duration} months)"
        self.ws["P1"].value = f"NEW\nNET PRICE\n({agreement_duration} months)"
        self.ws["Q1"].value = f"EXISTING\nPRORATED PRICE\n({remaining_duration} months)"
        self.ws["R1"].value = f"NEW\nPRORATED PRICE\n({remaining_duration} months)"
        self.ws["S1"].value = f"ESTIMATED\nCREDIT"
        self.ws["T1"].value = f"ESTIMATED\nINVOICE"
        self.ws["U1"].value = f"TRUE DELTA\nNET COST"

        for col in ["O", "P", "Q", "R"]:
            cell = self.ws[f"{col}1"]
            cell.alignment = Middle

    def bold_delta_rows(self, FirstRow, LastRow, LastColumn, delta_col="K"):
        for row in range(FirstRow, LastRow + 1):
            delta_value = self.ws[f"{delta_col}{row}"].value
            if delta_value and delta_value != 0:
                for col in range(1, LastColumn + 1):  # Columns A to J
                    cell = self.ws.cell(row=row, column=col)
                    cell.font = BodyBoldFont

    def drip(self):
        self.delta_deets_column_widths()
        self.delta_deets_number_formats()
        self.delta_deets_dynamic_headers()
        self.bold_delta_rows(start_row=2, end_row=self.ws.max_row)
        self.wb.save(self.filepath)

# ----------------------------SUMMARY TAB FORMATTING CLASS ----------------------------
class DeltaSummaryDrip:
    def __init__(self, filepath, sheet_name):
        self.filepath = filepath
        self.sheet_name = sheet_name
        self.wb = load_workbook(filepath)
        self.ws = self.wb[sheet_name]

        self.CurrencyColumns = {"UNIT LIST PRICE",	"UNIT NET PRICE", "ESTIMATED CREDIT", "ESTIMATED INVOICE",	"TRUE DELTA NET COST"}

    logo_path = r"C:\Users\mkb00\PROJECTS\PythonProjects\MKdelta\LOGO.jpg"

    def set_column_widths(self):
        for col in ["A", "B"]:
            self.ws.column_dimensions[col].width = 26
        for col in ["C", "D", "E", "F", "G"]:
            self.ws.column_dimensions[col].width = 11
        for col in ["H", "I", "J"]:
            self.ws.column_dimensions[col].width = 19
        self.ws.column_dimensions["K"].width = 1

    def general_info_drip(self):
        for row in range(3, 20):
            label_cell = self.ws[f"A{row}"]
            value_cell = self.ws[f"B{row}"]

            label_cell.fill = SubHeaderColor
            label_cell.font = SubHeaderFont
            label_cell.alignment = Lefty
            label_cell.border = Border(left=Border_SheThicc, right=Border_IttyBitty, top=Border_IttyBitty, bottom=Border_IttyBitty)

            value_cell.fill = BodyColor
            value_cell.font = BodyFont
            value_cell.alignment = Righty
            value_cell.border = Border(left=Border_IttyBitty, right=Border_SheThicc, top=Border_IttyBitty, bottom=Border_IttyBitty)

            if row in [10, 11, 12, 13]: # Apply date format to rows 10–13
                value_cell.number_format = "DD-MMM-YYYY"
            if row in [17, 18, 19]:# Apply currency format to rows 17-19
                value_cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE

        FirstLabelCell = self.ws["A3"]
        LastLabelCell = self.ws["A19"]
        FirstValueCell = self.ws["B3"]
        LastValueCell = self.ws["B19"]
        FirstLabelCell.border = SummaryGeneralInfoBorder1
        FirstValueCell.border = SummaryGeneralInfoBorder2
        LastLabelCell.border = SummaryGeneralInfoBorder3
        LastValueCell.border = SummaryGeneralInfoBorder4
        if row in [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]:
            self.ws.row_dimensions[row].height = 12
       
        self.ws.row_dimensions[20].height = 5

    def items_table_drip(self, FirstRow, LastRow):
        self.Leftys = {"SKU"}
        self.Middles = {"ALLOCATION STATUS", "EXISTING QTY", "NEW QTY", "DELTA QTY", "UNIT LIST PRICE", "UNIT NET PRICE"}
        self.Rightys = {"ESTIMATED CREDIT", "ESTIMATED INVOICE", "TRUE DELTA NET COST"}
        for row in range(FirstRow, LastRow + 1):
            self.ws.row_dimensions[row].height = 30 if row == FirstRow else 15
            for col in range(1, 11):  # Columns A to J
                cell = self.ws.cell(row=row, column=col)
                if row == FirstRow: # Header row style
                    cell.fill = HeaderColor
                    cell.font = HeaderFont
                    cell.alignment = Middle
                    cell.border = Border(
                        left=Border_IttyBitty,
                        right=Border_IttyBitty,
                        top=Border_IttyBitty,
                        bottom=Border_DoubleBottom
                    )
                else: # Item row style
                    cell.fill = BodyColor
                    cell.font = BodyFont
                    #cell.alignment = self.LeftAlign if col == 1 else self.CenterAlign
                    header_value = self.ws.cell(row=FirstRow, column=col).value
                    if header_value in self.Middles:
                        cell.alignment = Middle
                    elif header_value in self.Rightys:
                        cell.alignment = Righty
                    elif header_value in self.Leftys:
                        cell.alignment = Lefty
                
                    bottom_style = Border_DoubleBottom if row == LastRow-2 else Border_IttyBitty # Apply double bottom border to last row
                    cell.border = Border(
                        left=Border_IttyBitty,
                        right=Border_IttyBitty,
                        top=Border_IttyBitty,
                        bottom=bottom_style
                    )

    def totals_row_drip(self, row):
    # Merge A:G
        self.ws.merge_cells(FirstRow=row, FirstColumn=1, LastRow=row, LastCoumn=7)
        cell = self.ws.cell(row=row, column=1)
        cell.value = "TOTAL:"
        cell.font = SubHeaderFont
        cell.alignment = Lefty
        cell.fill = SubHeaderColor
        cell.border = Border(
        left=Border_IttyBitty,
        right=Border_IttyBitty,
        bottom=Border_DoubleBottom
        )
        # Format columns H–J
        for col in ["H", "I", "J"]:
            c = self.ws[f"{col}{row}"]
            c.font = SubHeaderFont
            c.alignment = Righty
            c.fill = SubHeaderColor
            c.border = Border(
                left=Border_IttyBitty,
                right=Border_IttyBitty,
                bottom=Border_SheThicc
            )
        self.ws.row_dimensions[row+1].height = 5

    def notes_drip(self, FirstRow):
        for i in range(4):  # Notes + 3 lines
            row = FirstRow + i
            # Apply thick border to all cells in A–J before merging
            for col in range(1, 11):  # Columns A to J
                cell = self.ws.cell(row=row, column=col)
                cell.border = Border(
                    left=Border_SheThicc if col == 1 else Border_IttyBitty,
                    right=Border_SheThicc if col == 10 else Border_IttyBitty,
                    top=Border_SheThicc if i == 0 else Border_IttyBitty,
                    bottom=Border_SheThicc if i == 3 else Border_IttyBitty
                )

            self.ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10) # Merge the cells

            anchor = self.ws.cell(row=row, column=1) # Style the anchor cell (A)
            anchor.alignment = Lefty
            anchor.fill = HeaderColor if i == 0 else SubHeaderColor
            anchor.font = HeaderFont if i == 0 else BodyFont

    def apply_outline_borders(self, item_end_row, notes_start_row):
        for row in self.ws.iter_rows(min_row=21, max_row=item_end_row + 1, min_col=1, max_col=10):
            for cell in row:
                self._outline(cell)

        for row in self.ws.iter_rows(min_row=notes_start_row, max_row=notes_start_row + 3, min_col=1, max_col=1):
            for cell in row:
                self._outline(cell)

    def _outline(self, cell):
        last_item_row = self.item_end_row
        totals_row = last_item_row
        notes_start_row = self.notes_start_row  # Store this in format()

        borders = {
            "left": Border_SheThicc if cell.column == 1 else Border_IttyBitty,
            "right": Border_SheThicc  if cell.column == 10 else Border_IttyBitty,
            "top": (
                Border_SheThicc  if cell.row in [3, 21, notes_start_row] else Border_IttyBitty
            ),
            "bottom": (
                Border_DoubleBottom if cell.row == last_item_row-1 else
                Border_SheThicc  if cell.row == totals_row else
                Border_IttyBitty
            )
        }
        #print(totals_row)
        cell.border = Border(**borders)

    def apply_summary_outline(self, last_notes_row):
        for row in range(2, last_notes_row + 1):
            for col in range(1, 11):  # Columns A to J
                cell = self.ws.cell(row=row, column=col)
                borders = {
                    "left": Border_SheThicc if col == 1 else Border_IttyBitty,
                    "right": Border_SheThicc if col == 10 else Border_IttyBitty,
                    "top": Border_SheThicc if row == 1 else Border_IttyBitty,
                    "bottom": Border_SheThicc if row == last_notes_row else Border_IttyBitty
                }
                cell.border = Border(**borders)
                
        for row in range(last_notes_row-3, last_notes_row + 1):
            for col in range(1, 11):  # Columns A to J
                cell = self.ws.cell(row=row, column=col)
                borders = {
                    "left": Border_SheThicc if col == 1 else Border_IttyBitty,
                    "right": Border_SheThicc if col == 1 else Border_IttyBitty,
                    "top": Border_SheThicc if row == 1 else Border_IttyBitty,
                    "bottom": Border_SheThicc if row == last_notes_row else Border_IttyBitty
                }
                cell.border = Border(**borders)

    def apply_currency_format(self, start_row, end_row): # Read headers from row 21 (the item header row)
        header = [self.ws.cell(row=start_row, column=col).value for col in range(1, 11)]
        for col_name in self.CurrencyColumns:
            if col_name in header:
                col_idx = header.index(col_name) + 1  # 1-based index
                for row in range(start_row + 1, end_row+2):  # Skip header row
                    cell = self.ws.cell(row=row, column=col_idx)
                    cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
       
    def White_Out(self):
        white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        # Define ranges to white out
        ranges = [
            ("A2", "J2"),
            ("C3", "J19"),
            ("A20", "J20")
        ]
        for start, end in ranges:
            for row in self.ws[start:end]:
                for cell in row:
                    cell.fill = white_fill

    def bold_delta_rows(self, start_row, end_row, delta_col="J", max_col=10):
        for row in range(start_row, end_row + 1):
            delta_value = self.ws[f"{delta_col}{row}"].value
            if delta_value and delta_value != 0:
                for col in range(1, max_col + 1):  # Columns A to J
                    cell = self.ws.cell(row=row, column=col)
                    cell.font = BodyBoldFont

    def insert_logo(self):
        try:
            img = Image(self.logo_path)
            img.width = 250  # Adjust as needed
            img.height = 110
            img.anchor = "I3"
            self.ws.add_image(img)
        except Exception as e:
            print(f"⚠️ Logo insertion failed: {e}")

    def format(self, item_end_row):
        self.item_end_row = self.ws.max_row - 5
        self.notes_start_row = item_end_row + 3  # Store for _outline
        self.set_column_widths()
        self.general_info_drip()
        self.items_table_drip(start_row=21, end_row=item_end_row)
        self.totals_row_drip(item_end_row + 1)
        self.notes_drip(item_end_row + 3)
        self.apply_outline_borders(item_end_row, item_end_row + 3)
        self.apply_summary_outline(last_notes_row=self.notes_start_row + 3)
        self.apply_currency_format(start_row=21, end_row=item_end_row)
        self.White_Out()
        self.insert_logo()
        self.wb.save(self.filepath)

