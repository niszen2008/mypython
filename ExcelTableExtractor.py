import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def parse_table_info(text_file_path):
    """
    Parse text file containing table and column information.
    Expected format:
    TableName1
    Column1, Column2, Column3
    
    TableName2
    Column1, Column2, Column3, Column4
    """
    tables = {}
    current_table = None
    
    with open(text_file_path, 'r') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            
            # If line doesn't contain comma, it's a table name
            if ',' not in line:
                current_table = line
                tables[current_table] = []
            else:
                # It's a list of columns
                columns = [col.strip() for col in line.split(',')]
                if current_table:
                    tables[current_table] = columns
    
    return tables

def create_excel_with_tables(tables_dict, output_file):
    """
    Create Excel file with separate sheets for each table.
    Each sheet will have the table name and column headers.
    """
    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)
    
    for table_name, columns in tables_dict.items():
        # Create new sheet for each table
        ws = wb.create_sheet(title=table_name[:31])  # Excel sheet names limited to 31 chars
        
        # Write table name in first row
        ws.append([f"Table: {table_name}"])
        
        # Write column headers in second row
        ws.append(columns)
        
        # Optional: Add some sample empty rows
        for i in range(5):
            ws.append([''] * len(columns))
    
    wb.save(output_file)
    print(f"Excel file created: {output_file}")

# Main execution
if __name__ == "__main__":
    # Input: Text file with table and column information
    input_text_file = "table_structure.txt"
    
    # Output: Excel file with separate sheets
    output_excel_file = "tables_output.xlsx"
    
    # Parse the text file
    tables = parse_table_info(input_text_file)
    
    print(f"Found {len(tables)} tables:")
    for table_name, columns in tables.items():
        print(f"  - {table_name}: {len(columns)} columns")
    
    # Create Excel file
    create_excel_with_tables(tables, output_excel_file)