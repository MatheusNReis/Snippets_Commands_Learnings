# Importing workbook and worksheet from file
import openpyxl
workbook = openpyxl.load_workbook(file_path)
worksheet = workbook[sheet_name]





Access active workbook and worksheet
import win32com.client # from pywin32 library
# Connect to the Excel application
excel = win32com.client.Dispatch("Excel.Application")
# Get the active workbook
workbook = excel.ActiveWorkbook
worksheet = workbook.worksheets(WorksheetName)





# From win32com.client Import (active worksheet)
# Function to get the last filled row in a specific column
def get_last_filled_row(worksheet, column):
    xlUp = -4162  # Numeric value for xlUp
    last_filled_row = worksheet.Cells(worksheet.Rows.Count, column).End(xlUp).Row
    return last_filled_row

# Get the last filled row in column
last_row = get_last_filled_row(worksheet, ColumnDescription) + 1 # +1 to adjust the merge cell reference 





# From win32com.client Import (active worksheet) 
# Function to get the top-left cell value if the cell is merged
def get_top_left_cell_value(worksheet, row, col):
    cell = worksheet.Cells(row, col)
    if cell.MergeCells:
        return worksheet.Cells(cell.MergeArea.Row, cell.MergeArea.Column).Value
    else:
        return cell.Value # Corresponding value of merged cell range






# Inline Breakpoint
import pdb # Import the pdb module
pdb.set_trace() # Put in the desired breakpoint line
# c-> continue, q-> quit






# Removes non-numeric values in DataFrame and create columns of sums
df = pd.DataFrame(data)
# Filter out rows that contain non-numeric values
numeric_df = df.apply(pd.to_numeric, errors='coerce') # pd.to_numeric: attempts to convert each element to a numeric type (e.g., integer or float
                                                        # errors='coerce': This parameter tells the function to convert any non-numeric values to NaN (Not a Number) instead of raising an error
# Drop rows with NaN values (which were non-numeric)
numeric_df = numeric_df.dropna() # DataFrame numeric_df that contain NaN values.
#Or replace NaN values
numeric_df = numeric_df.fillna(0) # Replace NaN values by 0 in dataframe
# Calculate the sum of the remaining rows
numeric_df['Sum'] = numeric_df.sum(axis=1) # Creates column called 'Sum' or update the existing column called 'Sum'
                                            # axis=1 means horizontal sum
print(numeric_df)





# Excel 200 columns labels list
excel_labels = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 'CA', 'CB', 'CC', 'CD', 'CE', 'CF', 'CG', 'CH', 'CI', 'CJ', 'CK', 'CL', 'CM', 'CN', 'CO', 'CP', 'CQ', 'CR', 'CS', 'CT', 'CU', 'CV', 'CW', 'CX', 'CY', 'CZ', 'DA', 'DB', 'DC', 'DD', 'DE', 'DF', 'DG', 'DH', 'DI', 'DJ', 'DK', 'DL', 'DM', 'DN', 'DO', 'DP', 'DQ', 'DR', 'DS', 'DT', 'DU', 'DV', 'DW', 'DX', 'DY', 'DZ', 'EA', 'EB', 'EC', 'ED', 'EE', 'EF', 'EG', 'EH', 'EI', 'EJ', 'EK', 'EL', 'EM', 'EN', 'EO', 'EP', 'EQ', 'ER', 'ES', 'ET', 'EU', 'EV', 'EW', 'EX', 'EY', 'EZ', 'FA', 'FB', 'FC', 'FD', 'FE', 'FF', 'FG', 'FH', 'FI', 'FJ', 'FK', 'FL', 'FM', 'FN', 'FO', 'FP', 'FQ', 'FR', 'FS', 'FT', 'FU', 'FV', 'FW', 'FX', 'FY', 'FZ', 'GA', 'GB', 'GC', 'GD', 'GE', 'GF', 'GG', 'GH', 'GI', 'GJ', 'GK', 'GL', 'GM', 'GN', 'GO', 'GP', 'GQ', 'GR']





# Apply Excel labels in dataframe columns
# Generate Excel-like column labels
num_columns = df.shape[1]
excel_labels = excel_labels[:num_columns] # num_columns means the number of columns to be labeled
df.columns = excel_labels





# Get the sheet name by index
sheet_name = pd.ExcelFile(file_path, engine='openpyxl').sheet_names[sheet_index_number]






# Create dataframe from specified worksheet
df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')





# Copy column data from a dataframe to a column of another dataframe
df['Total Acumulado'] = numeric_df['Total']




# Filter columns and getting its labels
columns_toshow_ = list(df.loc[:, 'A':'Z'].columns)
Or
columns_to_show = list(df.loc[:, 0:26].columns)





# Attribute values to lines with specific condition
df.loc[df['J'] == 'CONCLUÍDA', 'Total Acumulado'] = '100%'
# Explanation: in dataframe 'df', attribute value '100%' to lines of column labeled 'Total Acumulado' which have value 'CONCLUÍDA' in column labeled 'J'





# Create excel file from dataframe
df.to_excel(output_file_path, index=False) # index=False means not to print lines' indexes
