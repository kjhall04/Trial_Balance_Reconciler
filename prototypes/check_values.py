import pandas as pd

def parse_excel_by_header(file_path, target_column_name, sheet_name=0):
    """
    Parses an Excel file by dynamically finding the row containing the 
    target column name and using that as the header.
    """
    
    # 1. Load the first few rows into a temporary DataFrame (e.g., first 10 rows)
    #    with no header specified initially.
    temp_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=10)
    
    # 2. Iterate through the rows to find the header row index
    header_row_index = None
    for index, row in temp_df.iterrows():
        if target_column_name in row.values:
            header_row_index = index
            break
            
    if header_row_index is None:
        print(f"Target column '{target_column_name}' not found in the first 10 rows.")
        return None
        
    # 3. Reload the full data using the found header row index
    #    The 'header' parameter in read_excel uses the 0-based index of the row.
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row_index)
    
    print(f"Header found at row index: {header_row_index}. Data loaded successfully.")
    return df

# --- Example Usage ---
# Assuming you have a file named 'data_with_offset_header.xlsx' 
# where your actual column headers start on some row, and you want 
# to find the column 'Product Name'
file_path = 'Accounting_Project\\client tb.xlsx'
target_column = 'Debit' # Replace with your actual column name

df = parse_excel_by_header(file_path, target_column)

sum = 0

if df is not None:
    df_clean = df[~df.astype(str).apply(
        lambda row: row.str.contains("TOTAL", case=False, na=False)
    ).any(axis=1)]

    debit_total = pd.to_numeric(df_clean["Debit"], errors="coerce").sum()
    credit_total = pd.to_numeric(df_clean["Credit"], errors="coerce").sum()

    print("Debit total:", debit_total)
    print("Credit total:", credit_total)