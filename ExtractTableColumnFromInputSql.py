# This script extracts table names and column names from a complex Oracle SQL query file.
# It uses the reference Excel file (with tables as sheets and columns as headers) to validate and resolve columns.
# Assumptions:
# - Columns in the SQL are preferably qualified (e.g., table.column); unqualified columns are matched to any used table that has them in the reference.
# - Handles SELECT *, aliases, and basic resolution.
# - Output Excel has separate sheets for each used table, with extracted/used columns as headers (similar to the reference format).
# - Requires installation: pip install sql-metadata pandas openpyxl

import pandas as pd
from collections import defaultdict
from sql_metadata import Parser

def extract_tables_columns_from_sql(sql_file, reference_excel, output_file):
    """
    Extracts used table names and column names from an Oracle SQL query file,
    validates against the reference Excel, and creates a new Excel with separate sheets for each used table.
    """
    
    try:
        # Read the SQL query from file
        print(f"Reading SQL file: {sql_file}")
        with open(sql_file, 'r') as f:
            query = f.read().strip()
        
        # Parse the SQL query
        parser = Parser(query)
        raw_tables = parser.tables
        tables = [t.upper() for t in raw_tables]
        columns = parser.columns
        aliases = {k.upper(): v.upper() for k, v in parser.tables_aliases.items()}
        
        print(f"\nFound {len(tables)} tables in query: {', '.join(raw_tables)}")
        
        # Load reference Excel
        print(f"\nLoading reference: {reference_excel}")
        excel = pd.ExcelFile(reference_excel)
        table_to_columns = {}
        for sheet in excel.sheet_names:
            df = excel.parse(sheet, nrows=0)  # Read only headers
            table_to_columns[sheet.upper()] = [str(col).upper() for col in df.columns]
        
        # Collect used columns per table (use set for deduplication)
        used_columns = defaultdict(set)
        
        for col in columns:
            if '.' in col:
                # Qualified column (e.g., table.col or alias.col)
                parts = col.rsplit('.', 1)  # Split from right to handle db.schema.table.col if present, but assume simple
                if len(parts) == 2:
                    table_part, col_name = parts
                    table_part = table_part.upper()
                    col_name = col_name.upper()
                    
                    # Resolve if it's an alias
                    table = aliases.get(table_part, table_part)
                    
                    if table in tables:
                        if col_name == '*':
                            # Add all columns from reference for this table
                            if table in table_to_columns:
                                for c in table_to_columns[table]:
                                    used_columns[table].add(c)
                        else:
                            # Add specific column if it exists in reference
                            if table in table_to_columns and col_name in table_to_columns[table]:
                                used_columns[table].add(col_name)
                            else:
                                print(f"Warning: Column '{col}' not found in reference for table '{table}'")
            else:
                # Unqualified column (e.g., col) - match to any used table in reference that has it
                col_upper = col.upper()
                found = False
                for t in tables:
                    if t in table_to_columns and col_upper in table_to_columns[t]:
                        used_columns[t].add(col_upper)
                        found = True
                if not found:
                    print(f"Warning: Unqualified column '{col}' not found in any used table's reference")
        
        # Include tables even if no columns extracted (e.g., if only used in JOIN without selecting columns)
        for t in tables:
            if t not in used_columns and t in table_to_columns:
                used_columns[t] = set()  # Empty, but sheet will be created with no columns (or perhaps skip?)
        
        print(f"\nFound usage in {len(used_columns)} tables")
        
        # Create output Excel with separate sheets
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for table, cols in used_columns.items():
                # Sanitize sheet name
                sheet_name = table[:31]
                sheet_name = sheet_name.replace('/', '_').replace('\\', '_').replace('*', '_')
                sheet_name = sheet_name.replace('[', '_').replace(']', '_').replace(':', '_')
                sheet_name = sheet_name.replace('?', '_')
                
                # Create DataFrame with columns as headers (even if empty)
                col_list = sorted(list(cols))
                table_df = pd.DataFrame(columns=col_list)
                
                # Write to sheet
                table_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                print(f"  ✓ Created sheet: '{sheet_name}' with {len(col_list)} columns")
        
        print(f"\n✅ Success! Output saved to: {output_file}")
        
        # Print summary
        print("\nExtracted Usage Summary:")
        for table, cols in used_columns.items():
            print(f"\n  Table: {table}")
            print(f"  Columns: {', '.join(sorted(cols)) if cols else 'None extracted'}")
        
        return True
    
    except FileNotFoundError as e:
        print(f"❌ Error: File not found - {str(e)}")
        return False
    except Exception as e:
        print(f"❌ Error occurred: {str(e)}")
        return False

# Main execution
if __name__ == "__main__":
    
    # CONFIGURE THESE PATHS
    REFERENCE_EXCEL = "output_separated_tables.xlsx"  # The reference file from previous script
    INPUT_SQL = "complex_oracle_query.sql"            # Your input Oracle SQL query file path
    OUTPUT_EXCEL = INPUT_SQL.replace('.sql', '.xlsx') # Output Excel named based on input SQL file
    
    # Run the extraction
    print("=" * 60)
    print("SQL Query Table/Column Extractor")
    print("=" * 60)
    
    success = extract_tables_columns_from_sql(INPUT_SQL, REFERENCE_EXCEL, OUTPUT_EXCEL)
    
    if success:
        print("\n" + "=" * 60)
        print("Process completed successfully!")
        print("=" * 60)
    else:
        print("\n" + "=" * 60)
        print("Process failed. Please check the error messages above.")
        print("=" * 60)