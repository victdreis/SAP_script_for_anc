import pyodbc
import pandas as pd
 
def main():
    # List of accessible columns in MBEW (from your discovery)
    accessible_columns = [
        'MANDT', 'MATNR', 'BWKEY', 'BWTAR', 'LVORM', 'SALK3', 'VPRSV', 'VERPR',
        'STPRS', 'PEINH', 'BKLAS', 'VMKUM', 'VMSAL', 'VMVPR', 'VMVER', 'VMSTP',
        'VMPEI', 'VMBKL', 'VMSAV', 'VJKUM', 'VJSAL', 'VJVPR', 'VJVER', 'VJSTP',
        'VJPEI', 'VJBKL', 'VJSAV', 'LFGJA', 'LFMON', 'BWTTY', 'STPRV', 'LAEPR',
        'ZKPRS', 'ZKDAT', 'TIMESTAMP', 'BWPRS', 'BWPRH', 'VJBWS', 'VJBWH', 'VVJSL',
        'VVJLB', 'VVMLB', 'VVSAL', 'ZPLPR', 'ZPLP1', 'ZPLP2', 'ZPLP3', 'ZPLD1',
        'ZPLD2', 'ZPLD3', 'PPERZ', 'PPERL', 'PPERV', 'KALKZ', 'KALKL', 'KALKV',
        'KALSC', 'XLIFO', 'MYPOL', 'BWPH1', 'BWPS1', 'ABWKZ', 'PSTAT', 'KALN1',
        'KALNR', 'BWVA1', 'BWVA2', 'BWVA3', 'VERS1', 'VERS2', 'VERS3', 'HRKFT',
        'KOSGR', 'PPRDZ', 'PPRDL', 'PPRDV', 'PDATZ', 'PDATL', 'PDATV', 'EKALR',
        'VPLPR', 'MLMAA', 'MLAST', 'LPLPR', 'VKSAL', 'HKMAT', 'SPERW', 'KZIWL',
        'WLINL', 'ABCIW', 'BWSPA', 'LPLPX', 'VPLPX', 'FPLPX', 'LBWST', 'VBWST',
        'FBWST', 'EKLAS', 'QKLAS', 'MTUSE', 'MTORG', 'OWNPR', 'XBEWM', 'BWPEI',
        'MBRUE', 'OKLAS', 'DUMMY_VAL_INCL_EEW_PS', 'OIPPINV', 'OICURVAL',
        'OICURDATE', 'OICURQTY', 'OICURUT', 'OIFUTVAL', 'OIFUTDATE', 'OIFUTQTY',
        'OIFUTUT', 'OINVALQTY', 'OIREQUAT', 'OITAXKEY', 'OIHANTYP', 'OIHMTXGR'
    ]
   
    # Exclude the columns SALK3 and VKSAL.
    filtered_columns = [col for col in accessible_columns if col not in ('SALK3', 'VKSAL')]
   
    # Build the SELECT clause (wrap each column name in square brackets to handle special characters).
    query_columns = ", ".join(f'[{col}]' for col in filtered_columns)
    # Add a WHERE clause to filter rows where BWKEY (the valuation area) is one of the desired values.
    sql_query = f"SELECT {query_columns} FROM MBEW WHERE [BWKEY] IN ('2032','2096','2914')"
   
    print("Executing query:", sql_query)
   
    connection_string = "DSN=TDVCPBIP;dataSource=S10231;"
   
    try:
        conn = pyodbc.connect(connection_string)
        print("Connected using DSN 'TDVCPBIP' with dataSource=S10231.")
    except Exception as e:
        print("Connection error:", e)
        return
   
    try:
        # Execute the query and retrieve data into a DataFrame.
        df = pd.read_sql(sql_query, conn)
        print("The data has been retrieved from MBEW.")
        # print(df.head())
    except Exception as e:
        print("Error executing query:", e)
        conn.close()
        return
    finally:
        conn.close()
   
    # Save the result to an Excel file with formatting.
    output_file = "MBEW_filtered_query.xlsx"
    writer = pd.ExcelWriter(output_file, engine="xlsxwriter")
    df.to_excel(writer, sheet_name="Filtered", index=False)
   
    workbook = writer.book
    worksheet = writer.sheets["Filtered"]
    header_format = workbook.add_format({
        "bold": True,
        "text_wrap": True,
        "valign": "top",
        "fg_color": "#D7E4BC",
        "border": 1
    })
   
    # Auto-adjust column widths and format the header.
    for col_num, col_name in enumerate(df.columns):
        worksheet.write(0, col_num, col_name, header_format)
        max_width = max(df[col_name].astype(str).apply(len).max(), len(col_name))
        worksheet.set_column(col_num, col_num, max_width + 2)
   
    writer.close()
    print(f"Filtered query result saved as '{output_file}'.")
 
if __name__ == '__main__':
    main()
 