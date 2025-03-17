import pyodbc
import pandas as pd
import os
from datetime import datetime

# Define a function to format the NM value.
def format_nm(nm):
    # Remove leading zeros.
    s = str(nm).lstrip("0")
    # If the remaining string is too short, return it as is.
    if len(s) <= 6:
        return s
    prefix = s[:len(s)-6]
    infix = s[len(s)-6:len(s)-3]
    sufix = s[-3:]
    return f"{prefix}.{infix}.{sufix}"

def main():
    # Use the pre-configured DSN "TDVCPBIP" with an extra data source parameter.
    connection_string = "DSN=TDVCPBIP;dataSource=S10231;"
   
    try:
        conn = pyodbc.connect(connection_string)
        print("Connected using DSN 'TDVCPBIP' with dataSource=S10231.")
    except Exception as e:
        print("Connection error:", e)
        return

    # Define the SQL query for MARC table
    sql_marc = """
    SELECT
        WERKS,
        MATNR,
        DISGR,
        DISMM,
        MINBE,
        MABST,
        VSPVB,
        PLIFZ,
        LGRAD
    FROM MARC
    WHERE WERKS IN ('2032','2096','2914')
    """
   
    try:
        df_marc = pd.read_sql(sql_marc, conn)
        print("Data retrieved from MARC table.")
    except Exception as e:
        print("Error retrieving data from MARC:", e)
        conn.close()
        return
    
    # Define the SQL query for MBEW table
    sql_mbew = """
    SELECT
        BWKEY AS WERKS,
        MATNR,
        VERPR,
        LFMON
    FROM MBEW
    WHERE BWKEY IN ('2032','2096','2914')
    """
    
    try:
        df_mbew = pd.read_sql(sql_mbew, conn)
        print("Data retrieved from MBEW table.")
    except Exception as e:
        print("Error retrieving data from MBEW:", e)
        conn.close()
        return
    finally:
        conn.close()
    
    # Merge MARC and MBEW dataframes on WERKS and MATNR
    df = df_marc.merge(df_mbew, on=["WERKS", "MATNR"], how="left")
    
    # Rename the columns as per the transformation.
    df = df.rename(columns={
        "MATNR": "NM",
        "DISGR": "Grupo MRP",
        "DISMM":"Tipo de MRP",
        "MINBE": "PR",
        "MABST": "EM",
        "VSPVB": "SupM",
        "PLIFZ": "Lead time",
        "LGRAD": "Grau atend. (%)",
        # "VERPR": "Preço médio móvel",
        # "LFMON": "Mês de avaliação"
    })
    
    # Apply the custom NM formatting function to the "NM" column.
    df["NM"] = df["NM"].apply(format_nm)
   
    # Define output file path with today's date
    today_date = datetime.today().strftime('%Y-%m-%d')
    output_folder = "deliver"
    os.makedirs(output_folder, exist_ok=True)
    output_file = os.path.join(output_folder, f"parametros_de_ressuprimento_{today_date}.xlsx")
    
    # Save the result to an Excel file with formatting.
    writer = pd.ExcelWriter(output_file, engine="xlsxwriter")
    df.to_excel(writer, sheet_name="Transformed", index=False)
   
    workbook = writer.book
    worksheet = writer.sheets["Transformed"]
    header_format = workbook.add_format({
        "bold": True,
        "text_wrap": True,
        "valign": "top",
        "fg_color": "#D7E4BC",
        "border": 1
    })
   
    # Apply formatting to the header and auto-adjust column widths.
    for col_num, col_name in enumerate(df.columns):
        worksheet.write(0, col_num, col_name, header_format)
        max_width = max(df[col_name].astype(str).apply(len).max(), len(col_name))
        worksheet.set_column(col_num, col_num, max_width + 2)
   
    writer.close()  # Use close() instead of save() to finalize the file.
    print(f"Transformed MARC and MBEW data saved as '{output_file}'.")

if __name__ == '__main__':
    main()