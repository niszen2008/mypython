import pandas as pd
import re
from collections import defaultdict

def load_table_reference(reference_file):
    """
    Load table and column information from the reference Excel file.
    Returns a dictionary: {table_name: [list of columns]}
    """
    print(f"Loading table reference from: {reference_file}")
    
    table_reference = {}
    
    try:
        # Read all sheets from the reference Excel
        excel_file = pd.ExcelFile(reference_file)
        
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(reference_file, sheet_name=sheet_name)
            # Get column names from the sheet
            columns = df.columns.tolist()
            table_reference[sheet_name] = columns
            print(f"  ✓ Loaded table: {sheet_name} with {len(columns)} columns")
        
        print(f"\nTotal tables loaded: {len(table_reference)}")
        return table_reference
        
    except Exception as e:
        print(f"❌ Error loading reference file: {str(e)}")
        return {}


def read_sql_file(sql_file_path):
    """
    Read SQL query from file.
    """
    try:
        with open(sql_file_path, 'r', encoding='utf-8') as f:
            sql_content = f.read()
        print(f"✓ SQL file loaded: {sql_file_path}")
        return sql_content
    except Exception as e:
        print(f"❌ Error reading SQL file: {str(e)}")
        return ""


def extract_table_column_from_sql(sql_query, table_reference):
    """
    Extract table names and column names from SQL query using the reference.
    Returns a dictionary: {table_name: [list of columns found in query]}
    """
    results = defaultdict(list)
    
    # Convert SQL to uppercase for easier matching
    sql_upper = sql_query.upper()
    
    # Remove comments from SQL
    sql_upper = re.sub(r'--.*?$', '', sql_upper, flags=re.MULTILINE)
    sql_upper = re.sub(r'/\*.*?\*/', '', sql_upper, flags=re.DOTALL)
    
    print("\nAnalyzing SQL query...")
    
    # For each table in reference
    for table_name, columns in table_reference.items():
        table_upper = table_name.upper()
        
        # Check if table is referenced in SQL
        # Look for table name in FROM, JOIN clauses
        table_pattern = r'\b' + re.escape(table_upper) + r'\b'
                
        if re.search(table_pattern, sql_upper):
            print(f"\n  Found table: {table_name}")
            
            # For each column in this table, check if it's in the SQL
            for column in columns:
                column_upper = column.upper()
                
                # Look for column references (with or without table prefix)
                # Pattern 1: table.column
                pattern1 = r'\b' + re.escape(table_upper) + r'\.\s*' + re.escape(column_upper) + r'\b'
                # Pattern 2: just column name (if unique enough)
                pattern2 = r'\b' + re.escape(column_upper) + r'\b'
                
                if re.search(pattern1, sql_upper) or re.search(pattern2, sql_upper):
                    results[table_name].append(column)
                    print(f"    ✓ Column: {column}")
    
    return dict(results)


def save_results_to_excel(results, output_file):
    """
    Save extraction results to Excel file.
    Format: TableName | ColumnName
    """
    try:
        data = []
        
        for table_name, columns in results.items():
            for column in columns:
                data.append({
                    'TableName': table_name,
                    'ColumnName': column
                })
        
        if not data:
            print("\n⚠️  No tables/columns found in SQL query!")
            # Create empty DataFrame
            df = pd.DataFrame(columns=['TableName', 'ColumnName'])
        else:
            df = pd.DataFrame(data)
        
        # Save to Excel
        df.to_excel(output_file, index=False, sheet_name='SQL_Analysis')
        
        print(f"\n✅ Results saved to: {output_file}")
        print(f"   Total rows: {len(data)}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error saving results: {str(e)}")
        return False


def analyze_sql_query(reference_file, sql_file, output_file):
    """
    Main function to analyze SQL query and extract table/column information.
    """
    print("=" * 70)
    print("SQL Query Analyzer - Table & Column Extractor")
    print("=" * 70 + "\n")
    
    # Step 1: Load table reference
    table_reference = load_table_reference(reference_file)
    
    if not table_reference:
        print("❌ Failed to load table reference. Exiting.")
        return False
    
    # Step 2: Read SQL file
    sql_query = read_sql_file(sql_file)
    
    if not sql_query:
        print("❌ Failed to read SQL file. Exiting.")
        return False
    
    print(f"\nSQL Query length: {len(sql_query)} characters")
    
    # Step 3: Extract tables and columns
    results = extract_table_column_from_sql(sql_query, table_reference)
    
    # Step 4: Display summary
    print("\n" + "=" * 70)
    print("EXTRACTION SUMMARY")
    print("=" * 70)
    
    if results:
        for table_name, columns in results.items():
            print(f"\n📊 Table: {table_name}")
            print(f"   Columns found: {len(columns)}")
            for col in columns:
                print(f"     • {col}")
    else:
        print("\n⚠️  No matching tables or columns found in SQL query!")
    
    # Step 5: Save results
    print("\n" + "=" * 70)
    success = save_results_to_excel(results, output_file)
    
    if success:
        print("=" * 70)
        print("✅ Process completed successfully!")
        print("=" * 70)
    
    return success


# Main execution
if __name__ == "__main__":
    
    # CONFIGURE THESE PATHS
    REFERENCE_EXCEL = "output_separated_tables.xlsx"  # Table reference file
    SQL_INPUT_FILE = "complex_oracle_query.sql"                       # Your Oracle SQL file
    OUTPUT_EXCEL = "sql_analysis_result.xlsx"          # Output results file
    
    # Run the analysis
    analyze_sql_query(REFERENCE_EXCEL, SQL_INPUT_FILE, OUTPUT_EXCEL)
