# import pyodbc
# import pandas as pd

# # Define a function to format the NM value.
# def format_nm(nm):
#     # Remove leading zeros.
#     s = str(nm).lstrip("0")
#     # If the remaining string is too short, return it as is.
#     if len(s) <= 6:
#         return s
#     prefix = s[:len(s)-6]
#     infix = s[len(s)-6:len(s)-3]
#     sufix = s[-3:]
#     return f"{prefix}.{infix}.{sufix}"

# def main():
#     # Use the pre-configured DSN "TDVCPBIP" with an extra data source parameter.
#     connection_string = "DSN=TDVCPBIP;dataSource=S10231;"
   
#     try:
#         conn = pyodbc.connect(connection_string)
#         print("Connected using DSN 'TDVCPBIP' with dataSource=S10231.")
#     except Exception as e:
#         print("Connection error:", e)
#         return

#     # Define the SQL query with server-side filtering.
#     sql_marc = """
#     SELECT
#         WERKS,
#         MATNR,
#         DISGR,
#         DISMM,
#         DISLS,
#         MEINS,
#         MINBE,
#         MABST,
#         VSPVB,
#         PLIFZ,
#         LGRAD
#     FROM MARC
#     WHERE WERKS IN ('2032','2096','2914')
#     """
   
#     try:
#         df = pd.read_sql(sql_marc, conn)
#         print("Data retrieved from MARC table.")
#     except Exception as e:
#         print("Error retrieving data:", e)
#         conn.close()
#         return
#     finally:
#         conn.close()

#     # Rename the columns as per the transformation.
#     df = df.rename(columns={
#         "WERKS": "Centro",
#         "MATNR": "NM",
#         "DISGR": "Grupo MRP",
#         "DISMM": "Tipo de MRP",
#         "MEINS": "Unid.medida basica",
#         "DISLS": "RegraCalcTamLotes",
#         "MINBE": "PR",
#         "MABST": "EM",
#         "VSPVB": "SupM",
#         "PLIFZ": "Lead time",
#         "LGRAD": "Grau atend. (%)"
#     })

#     # Apply the custom NM formatting function to the "NM" column.
#     df["NM"] = df["NM"].apply(format_nm)
   
#     # Save the result to an Excel file with formatting.
#     output_file = "parametros_de_ressuprimento_MARC.xlsx"
#     writer = pd.ExcelWriter(output_file, engine="xlsxwriter")
#     df.to_excel(writer, sheet_name="Transformed", index=False)
   
#     workbook = writer.book
#     worksheet = writer.sheets["Transformed"]
#     header_format = workbook.add_format({
#         "bold": True,
#         "text_wrap": True,
#         "valign": "top",
#         "fg_color": "#D7E4BC",
#         "border": 1
#     })
   
#     # Apply formatting to the header and auto-adjust column widths.
#     for col_num, col_name in enumerate(df.columns):
#         worksheet.write(0, col_num, col_name, header_format)
#         max_width = max(df[col_name].astype(str).apply(len).max(), len(col_name))
#         worksheet.set_column(col_num, col_num, max_width + 2)

#     writer.close()  # Use close() instead of save() to finalize the file.
#     print(f"Transformed MARC data saved as '{output_file}'.")

# if __name__ == '__main__':
#     main()

import pyodbc
import pandas as pd

def format_nm(nm):
    s = str(nm).lstrip("0")
    if len(s) <= 6:
        return s
    prefix = s[:len(s)-6]
    infix = s[len(s)-6:len(s)-3]
    sufix = s[-3:]
    return f"{prefix}.{infix}.{sufix}"

def main():
    connection_string = "DSN=TDVCPBIP;dataSource=S10231;"
    
    try:
        conn = pyodbc.connect(connection_string)
        print("Connected using DSN 'TDVCPBIP' with dataSource=S10231.")
    except Exception as e:
        print("Connection error:", e)
        return
    
    sql_marc = """
    SELECT
        WERKS,
        MATNR,
        DISGR,
        DISMM,
        DISLS,
        MINBE,
        MABST,
        VSPVB,
        PLIFZ,
        LGRAD
    FROM MARC
    WHERE WERKS IN ('2032','2096','2914')
    """
    
    sql_mara = """
    SELECT
        MATNR,
        MEINS
    FROM MARA
    """
    
    try:
        df_marc = pd.read_sql(sql_marc, conn)
        df_mara = pd.read_sql(sql_mara, conn)
        print("Data retrieved from MARC and MARA tables.")
    except Exception as e:
        print("Error retrieving data:", e)
        conn.close()
        return
    finally:
        conn.close()
    
    # Merge MARC and MARA using WERKS and MATNR as keys
    df = df_marc.merge(df_mara, on='MATNR', how='left')
    
    # Rename columns
    df = df.rename(columns={
        "WERKS": "Centro",
        "MATNR": "NM",
        "DISGR": "Grupo MRP",
        "DISMM": "Tipo de MRP",
        "MEINS": "Unid.medida basica",
        "DISLS": "RegraCalcTamLotes",
        "MINBE": "PR",
        "MABST": "EM",
        "VSPVB": "SupM",
        "PLIFZ": "Lead time",
        "LGRAD": "Grau atend. (%)"
    })
    
    # Apply the NM formatting function
    df["NM"] = df["NM"].apply(format_nm)
    
    # Save to Excel
    output_file = "parametros_de_ressuprimento_MARC.xlsx"
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
    
    # Format header and adjust column width
    for col_num, col_name in enumerate(df.columns):
        worksheet.write(0, col_num, col_name, header_format)
        max_width = max(df[col_name].astype(str).apply(len).max(), len(col_name))
        worksheet.set_column(col_num, col_num, max_width + 2)
    
    writer.close()
    print(f"Transformed MARC data saved as '{output_file}'.")

if __name__ == '__main__':
    main()