import pandas as pd
from openpyxl import Workbook
import os
import win32com.client as win32
win32c = win32.constants
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo

# Load the data from the Excel file
file_path = "my_tests/Data Analysis 11-21-2024.xlsx"  # Replace with your file path
# sheet1 = pd.read_excel(file_path, sheet_name="FY 24-25 SharePoint")
active_caseload_df = pd.read_excel(file_path, sheet_name="Active Assignments")

def pivot_table(wb: object, 
                ws1: object, 
                pt_ws: object, 
                ws_name: str, 
                pt_name: str, 
                pt_rows: list, 
                pt_cols: list, 
                pt_filters: list, 
                pt_fields: list, 
                start_row: int = 1, 
                start_col: int = 1,
                apply_grouping: bool = False,
                pt_visible_rows: list = None):
    """
    wb = workbook1 reference
    ws1 = worksheet1
    pt_ws = pivot table worksheet number
    ws_name = pivot table worksheet name
    pt_name = name given to pivot table
    pt_rows, pt_cols, pt_filters, pt_fields: values selected for filling the pivot tables
    start_row = starting row number for pivot table (default 1)
    start_col = starting column number for pivot table (default 1)
    apply_grouping = boolean to apply grouping to date fields (default False)
    pt_visible_rows = list of rows to make visible in pivot table (default [])
    """

    
    
    
    
    
    
    # Add title above pivot table
    pt_ws.Range(pt_ws.Cells(start_row, start_col), pt_ws.Cells(start_row + 1, start_col + 1)).Merge()
    pt_ws.Cells(start_row, start_col).Value = pt_name
    pt_ws.Cells(start_row, start_col).Font.Bold = True
    pt_ws.Cells(start_row, start_col).Font.Size = 12
    pt_ws.Cells(start_row, start_col).Font.Name = "Aptos Narrow"
    pt_ws.Cells(start_row, start_col).Font.ColorIndex = win32c.xlThemeColorAccent6  # Green Accent 6
    pt_ws.Cells(start_row, start_col).HorizontalAlignment = win32c.xlCenter  # Center align
    pt_ws.Cells(start_row, start_col).VerticalAlignment = win32c.xlCenter  # Middle align
    pt_ws.Cells(start_row, start_col).WrapText = True  # Wrap text
    # Merge cells across the table width
    
    
    # pivot table location - add 1 to account for title
    pt_loc_row = start_row + len(pt_filters) + 2
    
    # grab the pivot table source data
    pc = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=ws1.UsedRange)
    
    # create the pivot table object
    pc.CreatePivotTable(TableDestination=f'{ws_name}!R{pt_loc_row}C{start_col}', TableName=pt_name)

    # selecte the pivot table work sheet and location to create the pivot table
    pt_ws.Select()
    pt_ws.Cells(pt_loc_row, start_col).Select()
    
    for field_list, field_r in ((pt_filters, win32c.xlPageField), (pt_rows, win32c.xlRowField), (pt_cols, win32c.xlColumnField)):
        for i, value in enumerate(field_list):
            pf = pt_ws.PivotTables(pt_name).PivotFields(value)
            pf.Orientation = field_r
            pf.Position = i + 1
            
            # Check if the field name contains date-related keywords
            if apply_grouping:
                if any(date_key in value.lower() for date_key in ['date', 'months', 'years']):
                    
                    try:
                    # Group by months, quarters, and years
                        pf.AutoGroup()  # Auto group the dates
                    except Exception as e:
                        print(e)
                        print(f"Could not group field: {value}")
                        
    for field in pt_ws.PivotTables(pt_name).VisibleFields:
        name = field.Name
        try:
            pf = pt_ws.PivotTables(pt_name).PivotFields(name)
            pf.Orientation = win32c.xlHidden 
        except Exception as e:
            print(e)
            print(f"Could not hide field: {name}")
            
    if pt_visible_rows is None:
        pt_visible_rows = pt_rows
        
        
    for x in pt_visible_rows:
        pf = pt_ws.PivotTables(pt_name).PivotFields(x)
        pf.Orientation = win32c.xlRowField 


    # Sets the Values of the pivot table
    for field in pt_fields:
        pt_ws.PivotTables(pt_name).AddDataField(pt_ws.PivotTables(pt_name).PivotFields(field[0]), field[1], field[2]).NumberFormat = field[3]

    # Visiblity True or Valse
    pt_ws.PivotTables(pt_name).ShowValuesRow = True
    pt_ws.PivotTables(pt_name).ColumnGrand = True
    
    # Add black borders to all cells in pivot table
    pt_range = pt_ws.PivotTables(pt_name).TableRange2
    pt_range.Borders.LineStyle = win32c.xlContinuous
    pt_range.Borders.Color = 0x000000  # Black
    pt_range.Borders.Weight = win32c.xlThin

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True  # False


base_dir = os.path.dirname(os.path.abspath(__file__))
filename = os.path.join(base_dir, 'Data Analysis 11-21-2024.xlsx')


wb = excel.Workbooks.Open(filename)

# List all worksheet names
print("\nWorksheet names:")
for sheet in wb.Sheets:
    print(f"- {sheet.Name}")

ws_active = wb.Sheets('Active Assignments')
ws_sharepoint = wb.Sheets('FY 24-25 SharePoint')
ws2_name = 'pivot_table'
wb.Sheets.Add().Name = ws2_name
ws2 = wb.Sheets(ws2_name)

pt_name = 'Active Caseload by Negotiator, Status, Agreement Type, and Delinquency'
pt_rows = ['Negotiator', 'Status', 'Agreement Type', 'Delinquency']
pt_cols = []
pt_filters = []
# pt_values = ['']
# 

# help(win32c)

pt_fields = [['#', 'Sum of #', win32c.xlSum, '0']]

pivot_table(wb, ws_active, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields, start_row=1, start_col=1)


pt_name = 'Executed Agreements by Negotiator'
pt_rows = ['Negotiator', 'FE Date', 'Date Assigned']# 'Months (FE Date)', 'Years (FE Date)']
pt_cols = []
pt_filters = []
pt_fields = [['#', 'Sum of #', win32c.xlSum, '0']]
visible_row_fields = ['Negotiator', 'FE Date']
pt_visible_rows = ['Negotiator', 'Years (FE Date)',  'Months (FE Date)', 'FE Date']
pivot_table(wb, ws_sharepoint, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields, start_row=1, start_col=4, apply_grouping=True, pt_visible_rows=pt_visible_rows)