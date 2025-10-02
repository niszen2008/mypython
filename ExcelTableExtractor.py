import pandas as pd
from collections import defaultdict

def extract_tables_to_separate_sheets(input_file, output_file):
    """
    Reads an Excel file where:
    - Column 1 contains table names (rowwise)
    - Column 2 contains column names (rowwise)
    
    Creates a new Excel file with:
    - One sheet per unique table name
    - Each sheet contains the column names for that table as headers
    """
    
    try:
        # Read the input Excel file (first sheet)
        print(f"Reading input file: {input_file}")
        df = pd.read_excel(input_file, sheet_name=0)
        
        # Get the first two columns (table names and column names)
        # Rename for clarity
        df = df.iloc[:, :2]  # Select only first 2 columns
        df.columns = ['TableName', 'ColumnName']
        
        # Remove rows with missing values
        df = df.dropna()
        
        # Group column names by table name
        table_columns = defaultdict(list)
        
        for index, row in df.iterrows():
            table_name = str(row['TableName']).strip()
            column_name = str(row['ColumnName']).strip()
            table_columns[table_name].append(column_name)
        
        print(f"\nFound {len(table_columns)} unique tables")
        
        # Create new Excel file with separate sheets
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            
            for table_name, columns in table_columns.items():
                # Sanitize sheet name (Excel has 31 char limit and special char restrictions)
                sheet_name = table_name[:31]
                sheet_name = sheet_name.replace('/', '_').replace('\\', '_').replace('*', '_')
                sheet_name = sheet_name.replace('[', '_').replace(']', '_').replace(':', '_')
                sheet_name = sheet_name.replace('?', '_')
                
                # Create a DataFrame with these columns as headers
                table_df = pd.DataFrame(columns=columns)
                
                # Optionally add empty rows for data entry (uncomment if needed)
                # for i in range(10):
                #     table_df.loc[i] = [''] * len(columns)
                
                # Write to Excel sheet
                table_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                print(f"  ✓ Created sheet: '{sheet_name}' with {len(columns)} columns")
        
        print(f"\n✅ Success! Output saved to: {output_file}")
        print("\nTable Summary:")
        for table_name, columns in table_columns.items():
            print(f"\n  Table: {table_name}")
            print(f"  Columns: {', '.join(columns)}")
        
        return True
        
    except FileNotFoundError:
        print(f"❌ Error: Input file '{input_file}' not found!")
        return False
    except Exception as e:
        print(f"❌ Error occurred: {str(e)}")
        return False


# Main execution
if __name__ == "__main__":
    
    # CONFIGURE THESE PATHS
    INPUT_EXCEL = "input_table_details.xlsx"     # Your input Excel file path
    OUTPUT_EXCEL = "output_separated_tables.xlsx" # Your output Excel file path
    
    # Run the extraction
    print("=" * 60)
    print("Excel Table Extractor")
    print("=" * 60)
    
    success = extract_tables_to_separate_sheets(INPUT_EXCEL, OUTPUT_EXCEL)
    
    if success:
        print("\n" + "=" * 60)
        print("Process completed successfully!")
        print("=" * 60)
    else:
        print("\n" + "=" * 60)
        print("Process failed. Please check the error messages above.")
        print("=" * 60)