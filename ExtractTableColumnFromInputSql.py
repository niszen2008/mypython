import pandas as pd
from collections import defaultdict
import re

def extract_tables_columns_from_sql(sql_file, reference_excel, output_file):
    """
    Extracts used table names and column names from an Oracle SQL query file,
    including inner SQL and WITH clauses, validates against the reference Excel,
    and creates a new Excel with separate sheets for each used table.
    Ignores WIDTH statements and outer XML details.
    """
    
    try:
        # Read the SQL query from file
        print(f"Reading SQL file: {sql_file}")
        with open(sql_file, 'r') as f:
            query = f.read().strip()
        
        # Clean the query: Remove WIDTH statements and XML-related content
        query = re.sub(r'\bCOLUMN\s+\w+\s+WIDTH\s+\d+\s*', '', query, flags=re.IGNORECASE)
        query = re.sub(r'\b(XMLTYPE|XMLAGG|XML\s*\([^)]*\))\b', '', query, flags=re.IGNORECASE)
        
        # Normalize query: Remove extra whitespace, newlines for easier parsing
        query = ' '.join(query.split())
        
        # Extract CTE names (WITH clause)
        cte_names = set()
        cte_pattern = r'\bWITH\s+((?:\w+\s*(?:,\s*\w+\s*)*)\s*AS\s*\([^)]*\))'
        cte_matches = re.finditer(cte_pattern, query, re.IGNORECASE)
        for match in cte_matches:
            cte_list = match.group(1).split(',')
            for cte in cte_list:
                cte_name = cte.strip().split()[0]
                cte_names.add(cte_name.upper())
        
        # Extract table names and aliases from FROM/JOIN clauses, including subqueries
        tables = set()
        aliases = {}
        # Pattern for FROM/JOIN, capturing tables and aliases, including subqueries
        table_pattern = r'\b(FROM|JOIN)\s+((?:[\w.]+|\(\s*SELECT\s+[^)]+\))\s*(?:AS\s+|\s+)(\w+)?'
        subquery_pattern = r'\(\s*SELECT\s+.*?FROM\s+([\w.]+)\s*(?:AS\s+|\s+)(\w+)?'
        
        # Main query tables
        for match in re.finditer(table_pattern, query, re.IGNORECASE):
            clause, table_ref, alias = match.groups()
            if table_ref.startswith('('):
                # Handle subquery
                for sub_match in re.finditer(subquery_pattern, table_ref, re.IGNORECASE):
                    table_name, sub_alias = sub_match.groups()
                    table_name = table_name.split('.')[-1].upper()
                    if table_name not in cte_names:
                        tables.add(table_name)
                        if sub_alias:
                            aliases[sub_alias.upper()] = table_name
            else:
                table_name = table_ref.split('.')[-1].upper()
                if table_name not in cte_names:
                    tables.add(table_name)
                    if alias:
                        aliases[alias.upper()] = table_name
        
        print(f"\nFound {len(tables)} tables in query: {', '.join(tables)}")
        
        # Extract column names (qualified and unqualified)
        columns = set()
        column_pattern = r'\b(?:(\w+)\.)?(\w+)\b(?!\s*(?:=|\(|,|\s+AS\s+|\s+FROM|\s+WHERE|\s+GROUP|\s+ORDER|\s+JOIN))'
        for match in re.finditer(column_pattern, query, re.IGNORECASE):
            table_part, col_name = match.groups()
            col_name = col_name.upper()
            if col_name not in ('SELECT', 'FROM', 'JOIN', 'WHERE', 'GROUP', 'ORDER', 'BY', 'AND', 'OR', 'AS', 'ON', '*'):
                if table_part:
                    columns.add(f"{table_part.upper()}.{col_name}")
                else:
                    columns.add(col_name)
        
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
                parts = col.rsplit('.', 1)
                if len(parts) == 2:
                    table_part, col_name = parts
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
                # Unqualified column - match to any used table in reference
                col_upper = col
                found = False
                for t in tables:
                    if t in table_to_columns and col_upper in table_to_columns[t]:
                        used_columns[t].add(col_upper)
                        found = True
                if not found:
                    print(f"Warning: Unqualified column '{col}' not found in any used table's reference")
        
        # Include tables even if no columns extracted (e.g., used in JOIN)
        for t in tables:
            if t not in used_columns and t in table_to_columns:
                used_columns[t] = set()  # Empty sheet for tables with no columns
        
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