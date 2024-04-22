import pandas as pd
import os
from tkinter import Tk, filedialog
from tkinter.filedialog import askopenfilename
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder

def get_year_week(date):
    """Returns the ISO year-week string from a date."""
    return f"{date.year}-{date.isocalendar().week:02d}"

def custom_xlookup(value, lookup_df, lookup_column, return_column):
    """Performs a lookup similar to Excel's XLOOKUP function."""
    match = lookup_df[lookup_df[lookup_column] == value]
    return match.iloc[0][return_column] if not match.empty else "N/A"

def read_file(file_path):
    """Reads a file based on its extension and returns a DataFrame."""
    if file_path.endswith('.csv'):
        return pd.read_csv(file_path, sep=';', header=None, dtype=str, skiprows=1)
    elif file_path.endswith('.xlsx'):
        return pd.read_excel(file_path)
    else:
        raise ValueError("Unsupported file format.")

def write_file(df, file_path, is_csv=True):
    """Writes a DataFrame to a file based on the specified format."""
    if is_csv:
        df.to_csv(file_path, index=False)
    else:
        df.to_excel(file_path, index=False)
    print(f"Data organized and saved to {file_path}")

def process_files(pingrm_file, codate_file):
    """Processes the given PINGRM and CoDate2-x files."""
    trim_and_filter_sheet(pingrm_file)
    import_and_organize_data(codate_file)

def trim_and_filter_sheet(pingrm_file):
    """For PINGRM schedule files, trims and filters the data."""
    print(f"Selected file: {pingrm_file}")



    if pingrm_file.endswith('.csv'):
        df = pd.read_csv(pingrm_file, sep=';', header=None, dtype=str, skiprows=1)
    elif pingrm_file.endswith('.xlsx'):
        df = pd.read_excel(pingrm_file)
    else:
        raise ValueError("Unsupported file format")

    print("Original DataFrame:")
    print(df.head())

    if pingrm_file.endswith('.csv'):
        df.replace(';;', ';', regex=True, inplace=True)
        print("DataFrame after replacing double semicolons:")
        print(df.head())

    processed_data = []

    for index, row in df.iterrows():
        print(f"Processing row {index + 1}")

        item_number = ""
        term_date = ""
        order_quantity = ""
        reference = ""
        year_week = ""

        for value in row:
            if pd.notna(value):
                if len(str(value)) == 10 and str(value).startswith('4'):
                    item_number = str(value)
                elif pd.to_datetime(value, errors='coerce') is not pd.NaT and term_date == "":
                    term_date = pd.to_datetime(value).strftime('%d-%m-%Y')
                    year_week = f"{pd.to_datetime(value).year}-{pd.to_datetime(value).isocalendar().week:02d}"
                elif str(value).isnumeric() and order_quantity == "" and term_date != "":
                    order_quantity = str(value)
                elif (len(str(value)) == 8 and str(value).startswith('3')) or str(value) == "forecast":
                    reference = str(value)

        print(f"Parsed data: Item Number: {item_number}, Term Date: {term_date}, Order Quantity: {order_quantity}, Reference: {reference}, Year Week: {year_week}")

        processed_data.append({
            'Item Number': item_number,
            'Term Date': term_date,
            'Order Quantity': order_quantity,
            'Reference': reference,
            'Year Week': year_week
        })

    df = df.iloc[1:]  # Delete the second row of the DataFrame

    processed_df = pd.DataFrame(processed_data)
    print("Processed DataFrame:")
    print(processed_df.head())

    if os.path.exists('output.xlsx'):
        excel_file = load_workbook('output.xlsx')
    else:
        excel_file = Workbook()
        excel_file.remove(excel_file.active)

    sheet = excel_file.create_sheet('Processed Data', index=0)
    for row in dataframe_to_rows(processed_df, index=False, header=True):
        sheet.append(row)

    # Set all columns to a specific width
    for column_cells in sheet.columns:
        width = 18
        for cell in column_cells:
            sheet.column_dimensions[cell.column_letter].width = width

    excel_file.save('output.xlsx')
    print(f"Processed data saved to 'output.xlsx'")

def import_and_organize_data(codate_file):
    """Cleans and prepares the order book for further processing."""
    # Read the CSV file, focusing only on the necessary columns
    df = pd.read_csv(codate_file, usecols=[
        'CustID', 'CustomerPONumber', 'CONumber', 'COLineNumber', 
        'ItemNumber', 'ItemDescription', 'CustomerItemNumber', 
        'ItemOrderedQuantity', 'OPEN_QTY', 'PromisedDlvryDate'
    ])
    
    # Renaming columns to match the expected names
    df.rename(columns={
        'COLineNumber': 'Ln',
        'ItemNumber': 'Item Number',
        'ItemDescription': 'Item Description',
        'CustomerItemNumber': 'Cust Item Number',
        'ItemOrderedQuantity': 'OrderQty',  # Updated based on CSV header
        'OPEN_QTY': 'Open Qty',  # Updated based on CSV header
        'PromisedDlvryDate': 'PromDlvry'  # Updated based on CSV header
    }, inplace=True)
    
    # Convert 'PromDlvry' to datetime format and compute 'Year-Week'
    try:
        df['PromDlvry'] = pd.to_datetime(df['PromDlvry'], format='%d-%b-%Y')
        df["Year-Week"] = df['PromDlvry'].apply(get_year_week)
    except Exception as e:
        print(f"Error converting 'PromDlvry' to datetime or computing 'Year-Week': {e}")
        return df

    # Initialize 'Match PING?' column with default values
    df['Match PING?'] = ''  # Placeholder for future data

    # Sort the DataFrame by 'Year-Week'
    df.sort_values('Year-Week', inplace=True)

    # Creating the filename for the output file
    output_filename = "organized_" + os.path.basename(codate_file).replace('.xlsx', '').replace('.csv', '') + "_new" + ('.xlsx' if codate_file.endswith('.xlsx') else '.csv')
    write_file(df, output_filename, is_csv=codate_file.endswith('.csv'))

def select_files():
    """Allows user to select PINGRM and CoDate2-x files for processing."""
    root = Tk()
    root.attributes('-fullscreen', True)  # Set the root window to fullscreen
    pingrm_file = filedialog.askopenfilename(title='Select PINGRM schedule file', filetypes=[('CSV Files', '*.csv'), ('Excel files', '*.xlsx')])
    codate_file = filedialog.askopenfilename(title='Select CoDate2-x file', filetypes=[('CSV Files', '*.csv'), ('Excel files', '*.xlsx')])

    if not pingrm_file or not codate_file:
        print("File selection cancelled.")
        return

    process_files(pingrm_file, codate_file)

if __name__ == "__main__":
    select_files()


