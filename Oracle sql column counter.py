import re
from collections import Counter

def extract_columns_from_query(sql_query):
    """
    Extract table columns and their counts from an Oracle SQL query.
    
    Args:
        sql_query (str): The SQL query to parse
        
    Returns:
        dict: Dictionary with column names as keys and their counts as values
    """
    # Remove comments
    sql_query = re.sub(r'--.*?$', '', sql_query, flags=re.MULTILINE)
    sql_query = re.sub(r'/\*.*?\*/', '', sql_query, flags=re.DOTALL)
    
    # Remove string literals to avoid false matches
    sql_query = re.sub(r"'[^']*'", "''", sql_query)
    
    # Convert to uppercase for easier parsing
    sql_upper = sql_query.upper()
    
    # List to store all column references
    columns = []
    
    # Pattern to match table.column or alias.column or just column
    # Matches: table.column, alias.column, or standalone column names
    column_pattern = r'\b([A-Z_][A-Z0-9_]*\.)?([A-Z_][A-Z0-9_]*)\b'
    
    # Find all potential column references
    matches = re.findall(column_pattern, sql_upper)
    
    # SQL keywords to exclude
    sql_keywords = {
        'SELECT', 'FROM', 'WHERE', 'AND', 'OR', 'NOT', 'IN', 'EXISTS',
        'JOIN', 'INNER', 'LEFT', 'RIGHT', 'OUTER', 'CROSS', 'FULL',
        'ON', 'AS', 'ORDER', 'BY', 'GROUP', 'HAVING', 'LIMIT',
        'INSERT', 'UPDATE', 'DELETE', 'CREATE', 'ALTER', 'DROP',
        'TABLE', 'VIEW', 'INDEX', 'DATABASE', 'SCHEMA',
        'IS', 'NULL', 'BETWEEN', 'LIKE', 'DISTINCT', 'ALL', 'ANY',
        'UNION', 'INTERSECT', 'MINUS', 'CASE', 'WHEN', 'THEN', 'ELSE', 'END',
        'ASC', 'DESC', 'INTO', 'VALUES', 'SET', 'ROWNUM', 'DUAL',
        'CONNECT', 'START', 'WITH', 'PRIOR', 'LEVEL', 'SYSDATE',
        'COUNT', 'SUM', 'AVG', 'MAX', 'MIN', 'DECODE', 'NVL', 'TO_DATE',
        'TO_CHAR', 'TO_NUMBER', 'TRUNC', 'ROUND', 'SUBSTR', 'LENGTH',
        'UPPER', 'LOWER', 'TRIM', 'LTRIM', 'RTRIM', 'REPLACE',
        'PARTITION', 'OVER', 'ROW_NUMBER', 'RANK', 'DENSE_RANK'
    }
    
    for prefix, column in matches:
        # Skip SQL keywords
        if column in sql_keywords:
            continue
            
        # Build full column name (with table/alias prefix if present)
        if prefix:
            full_column = f"{prefix.rstrip('.')}.{column}"
        else:
            full_column = column
            
        columns.append(full_column)
    
    # Count occurrences
    column_counts = Counter(columns)
    
    return dict(column_counts)


def display_results(column_counts):
    """
    Display the column counts in a formatted way.
    
    Args:
        column_counts (dict): Dictionary of column names and counts
    """
    if not column_counts:
        print("No columns found in the query.")
        return
    
    print("\n" + "="*60)
    print("COLUMN ANALYSIS RESULTS")
    print("="*60)
    print(f"{'Column Name':<40} {'Count':>10}")
    print("-"*60)
    
    # Sort by count (descending), then by name
    sorted_columns = sorted(column_counts.items(), 
                          key=lambda x: (-x[1], x[0]))
    
    for column, count in sorted_columns:
        print(f"{column:<40} {count:>10}")
    
    print("-"*60)
    print(f"{'TOTAL UNIQUE COLUMNS:':<40} {len(column_counts):>10}")
    print(f"{'TOTAL COLUMN REFERENCES:':<40} {sum(column_counts.values()):>10}")
    print("="*60 + "\n")


# Example usage
if __name__ == "__main__":
    # Example Oracle SQL query
    sample_query = """
    SELECT 
        e.employee_id,
        e.first_name,
        e.last_name,
        d.department_name,
        e.salary,
        e.hire_date
    FROM 
        employees e
        INNER JOIN departments d ON e.department_id = d.department_id
    WHERE 
        e.salary > 50000
        AND e.hire_date > TO_DATE('2020-01-01', 'YYYY-MM-DD')
        AND d.department_name IN ('IT', 'Sales')
    ORDER BY 
        e.salary DESC, e.last_name
    """
    
    print("Analyzing SQL Query:")
    print(sample_query)
    
    # Extract columns and their counts
    result = extract_columns_from_query(sample_query)
    
    # Display results
    display_results(result)
    
    # You can also use your own query
    print("\n" + "="*60)
    print("TO USE WITH YOUR OWN QUERY:")
    print("="*60)
    print("your_query = '''YOUR SQL QUERY HERE'''")
    print("result = extract_columns_from_query(your_query)")
    print("display_results(result)")