import argparse
from typing import List
import os
import logging
import traceback
from datetime import datetime
import re
import shutil
import time


import pandas as pd
import numpy as np
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

from autoexcel import config


logger = logging.getLogger(__name__)


def get_latest_file(directory, pattern):
    """
    Find the most recent file in directory matching the pattern.
    
    Args:
        directory (str): Directory path to search
        pattern (str): Regex pattern to match filenames
        
    Returns:
        str: Full path to the most recent file, or None if no files found
    """
    logger.debug(f"Searching for latest file in {directory} with pattern {pattern}")
    
    files = []
    for filename in os.listdir(directory):
        match = re.search(pattern, filename)
        if match:
            date_str = match.group(1)
            # Convert date string to datetime, handling both period and hyphen separators
            try:
                date = datetime.strptime(date_str.replace('.', '-'), '%m-%d-%Y')
                files.append((os.path.join(directory, filename), date))
            except ValueError as e:
                logger.warning(f"Couldn't parse date from filename {filename}: {e}")
    
    if not files:
        logger.warning(f"No matching files found in {directory}")
        return None
        
    # Sort by date and return the most recent file
    latest_file = max(files, key=lambda x: x[1])[0]
    logger.info(f"Found latest file: {latest_file}")
    return latest_file

def fy_analysis(raw_dir, processed_dir, assigned_date_filter=[datetime(2023, 7, 1), None]):
    """
    Process FY analysis using the most recent raw data file and template file.
    """
    os.makedirs(processed_dir, exist_ok=True)
    
    # Find most recent raw data and processed files
    raw_xlsx = get_latest_file(raw_dir, r'Raw Data (\d{1,2}[-\.]\d{1,2}[-\.]\d{4})\.xlsx')
   
    # Output filename using date from raw file
    raw_date = re.search(r'(\d{1,2}[-\.]\d{1,2}[-\.]\d{4})', os.path.basename(raw_xlsx)).group(1)
    formatted_date = raw_date.replace('.', '-')  # Standardize date format
    output_filename = os.path.join(processed_dir, f'Data Analysis {formatted_date}.xlsx')
    
    # Remove existing output file if it exists
    if os.path.exists(output_filename):
        logger.info(f"Removing existing output file: {output_filename}")
        os.remove(output_filename)
        
    # Find previous week's data analysis file
    previous_week_xlsx = get_latest_file(processed_dir, r'Data Analysis (\d{1,2}[-\.]\d{1,2}[-\.]\d{4})\.xlsx')
    
    if not raw_xlsx or not previous_week_xlsx:
        raise FileNotFoundError("Could not find required input files")
    
    # Read the data from the Excel file
    df = read_data(raw_xlsx)

    # Process the data
    df_processed, df_active_assignments = preprocess_data(df, assigned_date_filter)
    
    # Write the processed data to Excel
    write_output(df_processed, df_active_assignments, output_filename)

    # Copy worksheets

    fiscal_year, _ = get_current_fiscal_year()
    fy_analytics_ws_name= f'FY {str(fiscal_year-1)[-2:]}-{str(fiscal_year)[-2:]} Analytics'
    
    
    copy_excel_worksheet(previous_week_xlsx, output_filename, worksheet_names=[fy_analytics_ws_name, 'Caseload Analysis'])
    print('here')
    logger.info(f'Process completed. The processed data has been saved to {output_filename}')
    


def fy_analysis_from_template(raw_xlsx, template_xlsx, processed_dir='fy_analysis_processed', assigned_date_filter=[datetime(2023, 7, 1), None]):
    os.makedirs(processed_dir, exist_ok=True)
    
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
    fy_analytics_ws_name= f'FY {str(fiscal_year-1)[-2:]}-{str(fiscal_year)[-2:]} Analytics'
    copy_excel_worksheet(template_xlsx, output_filename, worksheet_names=[fy_analytics_ws_name, 'Caseload Analysis'])

    logger.info(f'Process completed. The processed data has been saved to {output_filename}')
    


def read_data(raw_xlsx):
    """Reads data from the raw Excel file."""
    df = pd.read_excel(raw_xlsx)
    return df

def preprocess_data(df, assigned_date_filter):
    """Processes the DataFrame according to specified steps."""
    logger.info('Preprocessing data.')
    df = df.copy()

    # Insert sequential numbers
    # df.insert(0, '#', range(1, len(df) + 1))
    df.insert(0, '#', [1] * len(df))

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
    df['Time to Execution'] = df.apply(lambda row: networkdays(row['Date Assigned'], row['FE Date']) if pd.notna(row['FE Date']) else None, axis=1)
    df["Today's Date"] = pd.to_datetime('today').normalize()
    df['Time Since Assignment'] = df.apply(lambda row: networkdays(row['Date Assigned'], row["Today's Date"]) 
                                         if not row['Status'] in ['Completed', 'Duplicate', 'Withdrawn'] else None, axis=1)

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

    logger.info('Preprocessing complete.')
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
    logger.info('Filtering assigned date.')
    df = df[df['Date Assigned'] >= start_fiscal_year]
    if end_fiscal_year:
        df = df[df['Date Assigned'] < end_fiscal_year]
    logger.info('Filtering assigned date complete.')
    return df

def networkdays(start_date, end_date):
    if pd.isnull(start_date) or pd.isnull(end_date):
        return None
    
    # Convert dates to datetime.date objects in correct format
    start = pd.to_datetime(start_date).date()
    end = pd.to_datetime(end_date).date()

    return np.busday_count(start, end) + 1  # Include end date



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
    
    logger.info('Writing output to Excel.')
    
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        fiscal_year, _ = get_current_fiscal_year()
        fy_sheet_name = f'FY {str(fiscal_year-1)[-2:]}-{str(fiscal_year)[-2:]} SharePoint'

        # Write data to worksheets
        df_original.to_excel(writer, index=False, sheet_name=fy_sheet_name)
        df_active_assignments.to_excel(writer, index=False, sheet_name="Active Assignments")

        # Apply formatting
        for sheet_name, df in zip([fy_sheet_name, "Active Assignments"], [df_original, df_active_assignments]):
            worksheet = writer.sheets[sheet_name]
            format_worksheet(worksheet, df)
            
    logger.info("Writing output to Excel complete.")
    
def format_worksheet(worksheet, df):
    """Applies formatting to the Excel worksheet."""
    logger.debug('Formatting worksheet.')
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
                
    logger.debug('Formatting worksheet complete.')

def copy_excel_worksheet(source_excel_path, target_excel_path, worksheet_names: List[str]):
    """Copies worksheets from the source Excel file to the target Excel file."""
    logger.info('Copying worksheets.')
    from win32com.client import Dispatch
    # print(dir(win32c))
    xl = Dispatch("Excel.Application")
    xl.Visible = False

    fiscal_year, _ = get_current_fiscal_year()
    
    logger.info(f"Opening {source_excel_path}")
    soruce_wb = xl.Workbooks.Open(Filename=os.path.abspath(source_excel_path))
    logger.info(f"Opening {target_excel_path}")
    target_wb = xl.Workbooks.Open(Filename=os.path.abspath(target_excel_path))
    logger.debug(f"Attempting to copy worksheets from {source_excel_path} to {target_excel_path}")
    try:
        for worksheet_name in worksheet_names:
            logger.info(f"Attempting to copy worksheet {worksheet_name}")
            source_ws = soruce_wb.Worksheets(worksheet_name)
            source_ws.Copy(Before=target_wb.Worksheets(1))
            logger.debug(f"Successfully copied worksheet {worksheet_name}")
            
    except Exception as e:
        logger.exception(e)
    finally:
        soruce_wb.Close(SaveChanges=False)
        target_wb.Close(SaveChanges=True)
        xl.Quit()

    logger.info('Copying worksheets complete.')
    
    
if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Run AutoExcel processing script.")
    
    data_dir=os.path.join(config.data_dir, 'fy_analysis')
    raw_dir = os.path.join(data_dir, 'raw')
    template_dir = os.path.join(data_dir, 'templates')
    
    
    # Add arguments
    parser.add_argument('--raw_dir', type=str, default=raw_dir, help="Relative path to the raw data directory.")
    parser.add_argument('--template_dir', type=str, default=template_dir, help="Relative path to the template directory.")
    parser.add_argument('--processed_dir', type=str, default='.', help="The output directory for the processed Excel file.")
    parser.add_argument('--data_dir', type=str, default=data_dir, help="Path to the data directory.")
    parser.add_argument('--start_date', type=str, default=None, help="Start date for filtering in YYYY-MM-DD format.")
    parser.add_argument('--end_date', type=str, default=None, help="End date for filtering in YYYY-MM-DD format.")

    # Parse arguments
    args = parser.parse_args()
    
    if args.start_date:
        print(*[int(x) for x in args.start_date.split('-')])
        start_date = datetime(*[int(x) for x in args.start_date.split('-')])
    
    end_date=None
    if args.end_date:
        end_date = datetime(*[int(x) for x in args.end_date.split('-')])
        
    assigned_date_filter=[start_date, end_date]

    fy_analysis(raw_dir, template_dir, processed_dir=args.processed_dir, assigned_date_filter=assigned_date_filter)