from typing import List
import os
import traceback
from datetime import datetime

import pandas as pd
import numpy as np
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

def main():
    # Parameters
    raw_xlsx = os.path.join('data', 'raw', 'Raw Data 9-19-2024.xlsx')
    processed_xlsx = os.path.join('data', 'templates', 'template_processed_workbook.xlsx')
    processed_dir = os.path.join('data', 'test')
    os.makedirs(processed_dir, exist_ok=True)
    assigned_date_filter = [datetime(2023, 7, 1), None]

    ################################################################################################
    # Script starts here
    ################################################################################################

    # Read the data from the Excel file
    df = read_data(raw_xlsx)

    # Process the data
    df_processed, df_active_assignments = preprocess_data(df, assigned_date_filter)

    # Output filename
    raw_date = os.path.basename(raw_xlsx).split('.')[0].split()[-1]
    output_filename = os.path.join(processed_dir, f'Data Analysis {raw_date}.xlsx')

    # Write the processed data to Excel
    write_output(df_processed, df_active_assignments, output_filename)

    # Copy worksheets

    fiscal_year, _ = get_current_fiscal_year()
    fy_analytics_ws_name= f'FY {str(fiscal_year)[-2:]}-{str(fiscal_year+1)[-2:]} Analytics'
    copy_excel_worksheet(processed_xlsx, output_filename, worksheet_names=[fy_analytics_ws_name, 'Caseload Analysis'])

    print('Process completed. The processed data has been saved to', output_filename)


def read_data(raw_xlsx):
    """Reads data from the raw Excel file."""
    df = pd.read_excel(raw_xlsx)
    return df

def preprocess_data(df, assigned_date_filter):
    """Processes the DataFrame according to specified steps."""
    df = df.copy()

    # Insert sequential numbers
    df.insert(0, '#', range(1, len(df) + 1))

    # Add new columns
    new_columns = ['Time to Assignment', 'Time to Execution', "Today's Date", 'Time Since Assignment', 'Delinquency']
    for col in new_columns:
        df[col] = None

    # Ensure date columns are in datetime format
    date_cols = ['Date Assigned', 'Date Received at OSP', 'FE Date']
    for col in date_cols:
        df[col] = pd.to_datetime(df[col])

    # Filter data based on assigned date
    df = filter_assigned_date(df, start_fiscal_year=assigned_date_filter[0], end_fiscal_year=assigned_date_filter[1])

    # Exclude certain 'Negotiator' values
    df = df[~df['Negotiator'].isin(['COE', 'NCE', 'OGC'])]

    # Calculate time metrics
    df['Time to Assignment'] = df.apply(lambda row: networkdays(row['Date Received at OSP'], row['Date Assigned']), axis=1)
    df['Time to Execution'] = df.apply(lambda row: networkdays(row['Date Assigned'], row['FE Date']), axis=1)
    df["Today's Date"] = pd.to_datetime('today').normalize()
    df['Time Since Assignment'] = df.apply(lambda row: networkdays(row['Date Assigned'], row["Today's Date"]), axis=1)

    # Format date columns
    date_cols_formatted = ['Date Assigned', 'Deadline Date', 'Date Received at OSP', 'Date Received at WVU', 'FE Date', "Today's Date"]
    for col in date_cols_formatted:
        if col in df.columns:
            df[col] = df[col].dt.strftime('%m/%d/%Y')

    # Sort and categorize
    df = df.sort_values(by='Time Since Assignment', ascending=False)
    df['Delinquency'] = df['Time Since Assignment'].apply(categorize_delinquency)

    # Create copies for different outputs
    df_original = df.copy()
    df_active_assignments = df_original.copy()
    df_active_assignments = df_active_assignments[~df_active_assignments['Status'].isin(['Completed', 'Duplicate', 'Withdrawn'])]

    return df_original, df_active_assignments

def get_current_fiscal_year():
    today = datetime.today()
    if today.month >= 10:
        fiscal_year_start = datetime(today.year, 10, 1)
        fiscal_year = today.year + 1
    else:
        fiscal_year_start = datetime(today.year - 1, 10, 1)
        fiscal_year = today.year
    return fiscal_year, fiscal_year_start

def filter_assigned_date(df, start_fiscal_year, end_fiscal_year):
    df = df[df['Date Assigned'] >= start_fiscal_year]
    if end_fiscal_year:
        df = df[df['Date Assigned'] < end_fiscal_year]
    return df

def networkdays(start_date, end_date):
    if pd.isnull(start_date) or pd.isnull(end_date):
        return None
    return np.busday_count(start_date.date(), end_date.date()) + 1  # Include end date

def categorize_delinquency(days):
    if pd.isnull(days):
        return None
    elif days >= 90:
        return '> 90 Days'
    elif days >= 60:
        return '> 60 Days'
    elif days >= 30:
        return '> 30 Days'
    else:
        return '< 30 Days'

def write_output(df_original, df_active_assignments, output_filename):
    """Writes the processed data to an Excel file with formatting."""
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        fiscal_year, _ = get_current_fiscal_year()
        fy_sheet_name = f'FY {str(fiscal_year)[-2:]} SharePoint'

        # Write data to worksheets
        df_original.to_excel(writer, index=False, sheet_name=fy_sheet_name)
        df_active_assignments.to_excel(writer, index=False, sheet_name="Active Assignments")

        # Apply formatting
        for sheet_name, df in zip([fy_sheet_name, "Active Assignments"], [df_original, df_active_assignments]):
            worksheet = writer.sheets[sheet_name]
            format_worksheet(worksheet, df)

def format_worksheet(worksheet, df):
    """Applies formatting to the Excel worksheet."""
    # Set default row height
    worksheet.sheet_format.defaultRowHeight = 15

    # Apply auto-filter
    max_column = get_column_letter(worksheet.max_column)
    worksheet.auto_filter.ref = f"A1:{max_column}1"

    # Format header
    title_fill = PatternFill(start_color="C0E6F4", end_color="C0E6F4", fill_type="solid")
    non_bold_font = Font(bold=False)
    for col in range(1, worksheet.max_column + 1):
        cell = worksheet.cell(row=1, column=col)
        cell.fill = title_fill
        cell.font = non_bold_font

    # Apply conditional formatting for 'Delinquency'
    red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
    orange_fill = PatternFill(start_color='FFFFA500', end_color='FFFFA500', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')

    if 'Delinquency' in df.columns:
        delinquency_col_idx = df.columns.get_loc('Delinquency') + 1
        delinquency_col_letter = get_column_letter(delinquency_col_idx)

        for idx, cell in enumerate(worksheet[delinquency_col_letter][1:], start=2):
            delinquency_value = cell.value
            if delinquency_value == '> 90 Days':
                cell.fill = red_fill
            elif delinquency_value == '> 60 Days':
                cell.fill = orange_fill
            elif delinquency_value == '> 30 Days':
                cell.fill = yellow_fill

def copy_excel_worksheet(source_excel_path, target_excel_path, worksheet_names: List[str]):
    """Copies worksheets from the source Excel file to the target Excel file."""
    from win32com.client import Dispatch
    xl = Dispatch("Excel.Application")
    xl.Visible = False

    fiscal_year, _ = get_current_fiscal_year()
    wb1 = xl.Workbooks.Open(Filename=os.path.abspath(source_excel_path))
    wb2 = xl.Workbooks.Open(Filename=os.path.abspath(target_excel_path))
    try:
        for worksheet_name in worksheet_names:
            print("Attempting to copy worksheet", worksheet_name)
            ws1 = wb1.Worksheets(worksheet_name)
            ws1.Copy(Before=wb2.Worksheets(1))

            if 'FY' in worksheet_name:
                ws2 = wb2.Worksheets(worksheet_name)
                ws2.Name = f'FY {str(fiscal_year)[-2:]}-{str(fiscal_year+1)[-2:]} Analytics'
    except Exception as e:
        print(e)
        traceback.print_exc()
    finally:
        wb1.Close(SaveChanges=False)
        wb2.Close(SaveChanges=True)
        xl.Quit()

if __name__ == '__main__':
    main()
