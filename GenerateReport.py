import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from datetime import datetime

###################
#FIELDS
####################
INPUT_DIR = #path to Sales report CV
OUTPUT_DIR = #path to outputted report
REPORT_NAME = #Report title
FILENAME_DATE_FORMAT = "%Y-%m-%d"

###################
#MAKE SALES REPORT METHODS
####################
#Make sure directory exists and if it doesn't, create one 
os.makedirs(OUTPUT_DIR, exist_ok=True)

#Load data from CV
def load_data(directory):
    """Load and concatenate all CSV files from the directory."""
    all_data = []
    for file in os.listdir(directory):
        if file.endswith(".csv"):
            df = pd.read_csv(os.path.join(directory, file))
            all_data.append(df)
    return pd.concat(all_data, ignore_index=True)
    
def transform_data(df):
    """Apply business logic to transform raw data."""
    df["Total"] = df["Quantity"] * df["Unit_Price"]
    df["Date"] = pd.to_datetime(df["Date"])
    df = df.sort_values(by="Date")
    return df
    
#Create Report
def write_excel_report(df, filename):
    """Write the DataFrame to a styled Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales Summary"

    # Write headers
    headers = list(df.columns)
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", fill_type="solid")

    ws.append(headers)
    for col_num, cell in enumerate(ws[1], 1):
        cell.font = header_font
        cell.fill = header_fill
        ws.column_dimensions[cell.column_letter].width = 15

    # Write rows
    for row in df.itertuples(index=False):
        ws.append(row)

    # Save
    output_path = os.path.join(OUTPUT_DIR, filename)
    wb.save(output_path)
    print(f"[✓] Report saved: {output_path}")

###################
#CREATE POWERPOINT
####################

###################
#MAIN
####################
