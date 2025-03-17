import pandas as pd

def read_excel_sheets(file_path):
    """
    Read all sheets from an Excel file and return them as a dictionary of DataFrames.
    
    Parameters:
    file_path (str): Path to the Excel file
    
    Returns:
    dict: Dictionary with sheet names as keys and pandas DataFrames as values
    """
    # Create Excel file object
    excel_file = pd.ExcelFile(file_path)
    
    # Get list of sheet names
    sheet_names = excel_file.sheet_names
    
    # Create a dictionary to store all sheets
    sheets_dict = {}
    
    # Read each sheet into a DataFrame
    for sheet in sheet_names:
        #sheets_dict[sheet] = pd.read_excel(file_path, sheet_name=sheet)
        sheets_dict[sheet] = pd.read_excel(file_path, sheet_name=sheet, skiprows=8, nrows=23)
        print(f"Read sheet: {sheet} with shape {sheets_dict[sheet].shape}")
    
    return sheets_dict

# Example usage
if __name__ == "__main__":
    # Replace with your Excel file path
    file_path = "D:\\Classwise 24 25 Sem I.xlsm"
    
    try:
        # Read all sheets
        all_sheets = read_excel_sheets(file_path)
        
        # Work with individual sheets
        for sheet_name, df in all_sheets.items():
            print(f"\nPreview of sheet: {sheet_name}")
            print(df)
            
            # Example operations
            print(f"Column names: {df.columns.tolist()}")
            print(f"Number of rows: {len(df)}")
            
    except FileNotFoundError:
        print("Error: Excel file not found!")
    except Exception as e:
        print(f"An error occurred: {str(e)}")