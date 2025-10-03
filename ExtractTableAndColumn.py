import re
#import cx_Oracle
import pandas as pd
from openpyxl import load_workbook
from collections import defaultdict

class OracleSQLAnalyzer:
    def __init__(self, connection_string):
        """
        Initialize with Oracle connection string
        Format: username/password@host:port/service_name
        """
        self.connection_string = connection_string
        self.conn = None
        
    def connect(self):
        """Establish Oracle database connection"""
        try:
            self.conn = cx_Oracle.connect(self.connection_string)
            print("Connected to Oracle database successfully")
        except Exception as e:
            print(f"Error connecting to database: {e}")
            raise
    
    def disconnect(self):
        """Close Oracle database connection"""
        if self.conn:
            self.conn.close()
            print("Disconnected from Oracle database")
    
    def extract_table_names(self, sql_query):
        """
        Extract table names from SQL query including WITH clauses
        """
        # Remove comments
        sql_query = re.sub(r'--.*$', '', sql_query, flags=re.MULTILINE)
        sql_query = re.sub(r'/\*.*?\*/', '', sql_query, flags=re.DOTALL)
        
        # Convert to uppercase for consistent matching
        sql_upper = sql_query.upper()
        
        tables = set()
        
        # Pattern to match table names after FROM and JOIN
        # Matches: FROM table_name, JOIN table_name
        from_pattern = r'\b(?:FROM|JOIN)\s+([A-Z_][A-Z0-9_]*\.)?([A-Z_][A-Z0-9_]*)\s*'
        
        matches = re.finditer(from_pattern, sql_upper)
        for match in matches:
            schema = match.group(1)
            table = match.group(2)
            
            # Skip CTE names by checking if they're defined in WITH clause
            if not self._is_cte_name(sql_upper, table):
                if schema:
                    tables.add(f"{schema.rstrip('.')}.{table}")
                else:
                    tables.add(table)
        
        return list(tables)
    
    def _is_cte_name(self, sql_upper, table_name):
        """Check if a name is a CTE (Common Table Expression)"""
        # Look for WITH clause definitions
        cte_pattern = r'\bWITH\s+.*?\b' + table_name + r'\s+AS\s*\('
        return bool(re.search(cte_pattern, sql_upper, re.DOTALL))
    
    def get_table_columns(self, table_name):
        """
        Execute DESC equivalent to get column details
        Returns: DataFrame with column details
        """
        if not self.conn:
            raise Exception("Not connected to database")
        
        cursor = self.conn.cursor()
        
        try:
            # Query to get column information (equivalent to DESC)
            query = """
                SELECT 
                    COLUMN_NAME,
                    DATA_TYPE,
                    DATA_LENGTH,
                    DATA_PRECISION,
                    DATA_SCALE,
                    NULLABLE
                FROM ALL_TAB_COLUMNS
                WHERE TABLE_NAME = :table_name
                ORDER BY COLUMN_ID
            """
            
            # Extract just table name if schema.table format
            parts = table_name.split('.')
            table_only = parts[-1] if len(parts) > 1 else table_name
            
            cursor.execute(query, {'table_name': table_only.upper()})
            
            columns = []
            for row in cursor:
                columns.append({
                    'Column Name': row[0],
                    'Data Type': row[1],
                    'Length': row[2],
                    'Precision': row[3],
                    'Scale': row[4],
                    'Nullable': row[5]
                })
            
            return pd.DataFrame(columns)
            
        except Exception as e:
            print(f"Error getting columns for {table_name}: {e}")
            return pd.DataFrame()
        finally:
            cursor.close()
    
    def create_reference_excel(self, sql_query, output_file='table_reference.xlsx'):
        """
        Create Excel file with table names as sheets and column details
        """
        tables = self.extract_table_names(sql_query)
        
        if not tables:
            print("No tables found in the SQL query")
            return
        
        print(f"Found {len(tables)} tables: {tables}")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for table in tables:
                print(f"Processing table: {table}")
                df = self.get_table_columns(table)
                
                if not df.empty:
                    # Excel sheet names have max 31 characters
                    sheet_name = table[:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"  Added {len(df)} columns to sheet '{sheet_name}'")
                else:
                    print(f"  No columns found for {table}")
        
        print(f"\nReference Excel file created: {output_file}")
    
    def analyze_query_with_reference(self, sql_query, reference_excel='table_reference.xlsx'):
        """
        Analyze SQL query and provide table-column mapping using reference Excel
        """
        # Load reference data
        excel_file = pd.ExcelFile(reference_excel)
        reference_data = {}
        
        for sheet in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet)
            reference_data[sheet] = df['Column Name'].tolist()
        
        # Extract tables from query
        tables_in_query = self.extract_table_names(sql_query)
        
        # Extract column references from query
        sql_upper = sql_query.upper()
        # Pattern to match column references (table.column or just column)
        column_pattern = r'\b([A-Z_][A-Z0-9_]*\.)?([A-Z_][A-Z0-9_]*)\b'
        
        potential_columns = set()
        for match in re.finditer(column_pattern, sql_upper):
            if match.group(2) not in ['SELECT', 'FROM', 'WHERE', 'JOIN', 'AND', 'OR', 'ON', 'AS', 'WITH']:
                potential_columns.add(match.group(0))
        
        # Match columns with tables
        results = defaultdict(list)
        
        for table in tables_in_query:
            table_key = table[:31]  # Match sheet name truncation
            if table_key in reference_data:
                for col in potential_columns:
                    # Extract column name without table prefix
                    col_name = col.split('.')[-1]
                    if col_name in reference_data[table_key]:
                        results[table].append(col_name)
        
        return dict(results)


# Example usage
def main():
    # Configuration
    CONNECTION_STRING = "username/password@host:port/service_name"
    
    # Read your complex SQL query
    with open('complex_query.txt', 'r') as f:
        sql_query = f.read()
    
    # Initialize analyzer
    analyzer = OracleSQLAnalyzer(CONNECTION_STRING)
    
    try:
        # Connect to database
        #analyzer.connect()
        
        # Step 1: Create reference Excel with table and column details
        print("=" * 60)
        print("Step 1: Creating reference Excel file...")
        print("=" * 60)
        analyzer.create_reference_excel(sql_query, 'table_reference.xlsx')
        
        # Step 2: Analyze query using reference
        print("\n" + "=" * 60)
        print("Step 2: Analyzing query with reference data...")
        print("=" * 60)
        results = analyzer.analyze_query_with_reference(sql_query, 'table_reference.xlsx')
        
        # Display results
        print("\nTable and Column Mapping:")
        print("-" * 60)
        for table, columns in results.items():
            print(f"\nTable: {table}")
            print(f"Columns: {', '.join(columns)}")
        
    finally:
        # Disconnect
        analyzer.disconnect()


if __name__ == "__main__":
    main()
