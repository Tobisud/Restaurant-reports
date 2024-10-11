import pandas as pd
import os
import re
from datetime import datetime


def view_report(file_path):
    xlsx_files = [f for f in os.listdir(file_path) if f.endswith('.xlsx')]
    # Read and print the contents of each .xlsx file
    for file_name in xlsx_files:
        file_path = os.path.join(file_path, file_name)
        
        # Load the Excel file
        excel_data = pd.ExcelFile(file_path)
        
        # Print sheet names
        print(f"File: {file_name}")
        print("Sheet names:", excel_data.sheet_names)
        
        # Load and print data from each sheet
        for sheet_name in excel_data.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            print(f"\nSheet: {sheet_name}")
            print(df)  # Display the first few rows of the dataframe
        print()  # Print a blank line for better readability
    
def get_saleWOTip(xlsx_file):
    sheet_name="Revenue center summary"
    df=pd.read_excel(xlsx_file,sheet_name=sheet_name)
     # Check if necessary columns exist
    if 'Revenue center' not in df.columns or 'Tax amount' not in df.columns or 'Net sales' not in df.columns :
        raise ValueError("The required columns are not present in the sheet")
    # Find out net sales
    net_sales = df[df['Revenue center'].str.lower() == 'dining room']['Net sales'].sum()
    tax_amount=df[df['Revenue center'].str.lower()=='dining room']['Tax amount'].sum()
    total_sales=round(net_sales+tax_amount,2)
    return total_sales

def get_date (xlsx_file):
    file_name=os.path.basename(xlsx_file)  
    # Define the regular expression pattern
    pattern = r'SalesSummary_(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})\.xlsx'
    match = re.match(pattern, file_name)
    if not match:
        raise ValueError(f"Filename '{file_name}' does not match the expected format")
    begin_date_str, end_date_str = match.groups()
    # Convert date strings to datetime objects
    try:
        begin_date = datetime.strptime(begin_date_str, '%Y-%m-%d')
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d')
    except ValueError as e:
        raise ValueError(f"Error parsing dates from filename: {e}")
    # Check if dates are the same
    if begin_date == end_date:
        # Format the date as M/D/YYYY
        return begin_date.strftime('%Y-%m-%d')  # Correct format for most systems
    else:
        raise ValueError(f"Begin date '{begin_date_str}' and end date '{end_date_str}' are not the same")

def get_location (xlsx_file):
    df = pd.read_excel(xlsx_file, sheet_name="All data")
    # Access the cell value in the second row, first column
    cell_value = df.iloc[0, 0]  # pandas is 0-indexed, so row 2, column 1 is index (1, 0)\
    #Check if the cell value is a string
    if isinstance(cell_value, str):
        # Extract the location from the string
        # Example input: "-Chin's Szechwan - Carlsbad"
        parts = cell_value.split('-')
        if len(parts) >= 2:
            location = parts[-1].strip()  # Get the last part and remove leading/trailing whitespace
            return location
        else:
            raise ValueError("Expected format not found in cell value")
    else:
        raise ValueError("Cell value is not a string")

def get_credit(xlsx_file):
    sheet_name="Payments summary"
    df=pd.read_excel(xlsx_file,sheet_name=sheet_name)
     # Check if necessary columns exist
    if 'Payment type' not in df.columns or 'Amount' not in df.columns:
        raise ValueError("The required columns are not present in the sheet")
    credit_debit_rows = df[df['Payment type'].str.lower() == 'credit/debit']
    credit_amount = credit_debit_rows.iloc[0]['Amount']
    return credit_amount

def get_gc(xlsx_file):
    sheet_name="Payments summary"
    df=pd.read_excel(xlsx_file, sheet_name=sheet_name)
     # Check if necessary columns exist
    if 'Payment type' not in df.columns or 'Amount' not in df.columns:
        raise ValueError("The required columns are not present in the sheet")
    gc_amount=df[df['Payment type'].str.lower()=='gift card']['Amount']
    return gc_amount.sum()

def get_cctips(xlsx_file):
    sheet_name="Payments summary"
    df=pd.read_excel(xlsx_file,sheet_name=sheet_name)
     # Check if necessary columns exist
    if 'Payment type' not in df.columns or 'Amount' not in df.columns:
        raise ValueError("The required columns are not present in the sheet")
    credit_debit_rows = df[df['Payment type'].str.lower() == 'credit/debit']
    cc_tips = credit_debit_rows.iloc[0]['Tips']
    return cc_tips

def get_cash(xlsx_file):
    sheet_name="Payments summary"
    df=pd.read_excel(xlsx_file,sheet_name=sheet_name)
     # Check if necessary columns exist
    if 'Payment type' not in df.columns or 'Amount' not in df.columns:
        raise ValueError("The required columns are not present in the sheet")
    cash_amounts = df[df['Payment type'].str.lower() == 'cash']['Amount']
    return cash_amounts.sum()

def get_pettycash(xlsx_file):
    sheet_name="Cash activity"
    df=pd.read_excel(xlsx_file,sheet_name=sheet_name)
     # Check if necessary columns exist
    if 'Cash adjustments' not in df.columns:
        raise ValueError("The required columns are not present in the sheet")
    pettycash_amounts = df['Cash adjustments']
    return pettycash_amounts.sum()

def get_dd(xlsx_file):
    sheet_name="Payments summary"
    df=pd.read_excel(xlsx_file,sheet_name=sheet_name)
     # Check if necessary columns exist
    if 'Payment sub type' not in df.columns or 'Amount' not in df.columns:
        raise ValueError("The required columns are not present in the sheet")
    dd_total=df[df['Payment sub type'].str.lower() == 'doordash']['Amount']
    dd_tax=df[df['Payment sub type'].str.lower()=='doordash']['Tax amount']
    dd_amount=dd_total-dd_tax
    return round(dd_amount.sum(),2)

def get_uber(xlsx_file):
    sheet_name="Payments summary"
    df=pd.read_excel(xlsx_file,sheet_name=sheet_name)
     # Check if necessary columns exist
    if 'Payment sub type' not in df.columns or 'Amount' not in df.columns:
        raise ValueError("The required columns are not present in the sheet")
    ub_total=df[df['Payment sub type'].str.lower() == 'uber eats']['Amount']
    ub_tax=df[df['Payment sub type'].str.lower()=='uber eats']['Tax amount']
    ub_amount=ub_total-ub_tax
    return round(ub_amount.sum(),2)

def get_gh(xlsx_file):
    sheet_name="Payments summary"
    df=pd.read_excel(xlsx_file,sheet_name=sheet_name)
     # Check if necessary columns exist
    if 'Payment sub type' not in df.columns or 'Amount' not in df.columns:
        raise ValueError("The required columns are not present in the sheet")
    gh_total=df[df['Payment sub type'].str.lower() == 'grubhub']['Amount']
    gh_tax=df[df['Payment sub type'].str.lower()=='grubhub']['Tax amount']
    gh_amount=gh_total-gh_tax
    return round(gh_amount.sum(),2)

def get_gh_Tip(xlsx_file):
    sheet_name="Payments summary"
    df=pd.read_excel(xlsx_file,sheet_name=sheet_name)
     # Check if necessary columns exist
    if 'Payment sub type' not in df.columns or 'Tips' not in df.columns:
        raise ValueError("The required columns are not present in the sheet")
    gh_tip=df[df['Payment sub type'].str.lower() == 'grubhub']['Tips']
    return gh_tip.sum()
def get_ub_Tip(xlsx_file):
    sheet_name="Payments summary"
    df=pd.read_excel(xlsx_file,sheet_name=sheet_name)
     # Check if necessary columns exist
    if 'Payment sub type' not in df.columns or 'Tips' not in df.columns:
        raise ValueError("The required columns are not present in the sheet")
    ub_tip=df[df['Payment sub type'].str.lower() == 'uber eats']['Tips']
    return ub_tip.sum()
def get_dd_Tip(xlsx_file):
    sheet_name="Payments summary"
    df=pd.read_excel(xlsx_file,sheet_name=sheet_name)
     # Check if necessary columns exist
    if 'Payment sub type' not in df.columns or 'Tips' not in df.columns:
        raise ValueError("The required columns are not present in the sheet")
    dd_tip=df[df['Payment sub type'].str.lower() == 'doordash']['Tips']
    return dd_tip.sum()
def get_def_amount(xlsx_file):
    sheet_name="Deferred summary" 
    df=pd.read_excel(xlsx_file,sheet_name=sheet_name)
    # Check if necessary columns exist
    if 'Deferred type' not in df.columns or 'Gross amount' not in df.columns:
        return 0
    else:
        def_amount=df[df['Deferred type'].str.lower()=='deferred (gift cards)']['Gross amount']
    return 0-def_amount.sum()
#print(get_gc(r'C:\Users\Thanh Lu\Desktop\salereport\OS\SalesSummary_2024-08-01_2024-08-31.xlsx'))




















