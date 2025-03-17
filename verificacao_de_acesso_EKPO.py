import pyodbc
import time
from concurrent.futures import ThreadPoolExecutor, TimeoutError
 
def get_table_columns(conn, table):
    """
    Retrieves a list of candidate column names for the given table using ODBC metadata.
    """
    cursor = conn.cursor()
    columns = []
    for row in cursor.columns(table=table):
        columns.append(row.column_name)
    return columns
 
def test_column_access(column, connection_string):
    """
    Attempts to query TOP 1 for a given column from EKPO.
    Returns a tuple (column, success, elapsed_time).
    """
    start = time.time()
    try:
        conn = pyodbc.connect(connection_string)
        cursor = conn.cursor()
        # Use square brackets to safely delimit column names.
        query = f"SELECT TOP 1 [{column}] FROM MARA"
        cursor.execute(query)
        cursor.fetchone()
        elapsed = time.time() - start
        conn.close()
        return (column, True, elapsed)
    except Exception as e:
        try:
            conn.close()
        except Exception:
            pass
        return (column, False, None)
 
def main():
    connection_string = "DSN=TDVCPBIP;dataSource=S10231;"
    try:
        conn = pyodbc.connect(connection_string)
        print("Connected using DSN 'TDVCPBIP' with dataSource=S10231.")
    except Exception as e:
        print("Connection error:", e)
        return
 
    # Discover candidate columns in EKPO.
    candidate_columns = get_table_columns(conn, "EKPO")
    print("Discovered candidate columns in EKPO:")
    print(candidate_columns)
    conn.close()
    accessible = []
    inaccessible = []
    # Use a ThreadPoolExecutor to test each candidate column with a 5-second timeout.
    with ThreadPoolExecutor(max_workers=5) as executor:
        future_to_column = {executor.submit(test_column_access, col, connection_string): col 
                            for col in candidate_columns}
        for future in future_to_column:
            col = future_to_column[future]
            try:
                result = future.result(timeout=5)
                # result is a tuple: (column, success, elapsed_time)
                if result[1]:
                    accessible.append(result)
                else:
                    inaccessible.append(col)
            except TimeoutError:
                print(f"Column '{col}' timed out.")
                inaccessible.append(col)
            except Exception as e:
                print(f"Column '{col}' raised exception: {e}")
                inaccessible.append(col)
    print("\nAccessible columns in EKPO (with access times in seconds):")
    for col, success, elapsed in accessible:
        print(f"{col}: {elapsed:.2f} seconds")
    print("\nInaccessible columns in EKPO:")
    print(inaccessible)
 
if __name__ == '__main__':
    main()