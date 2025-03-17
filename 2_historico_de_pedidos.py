import win32com.client
import time
import logging
import os
import polars as pl
import pandas as pd

# Configure logging to write to a file and the console.
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler("sap_query.log"),
        logging.StreamHandler()
    ]
)

class SAPAutomation:

    def connect_to_sap(self):
        try:
            sap_gui = win32com.client.GetObject("SAPGUI")
            application = sap_gui.GetScriptingEngine
            session = application.Children(0).Children(0)
            logging.info("Connected to SAP session successfully.")
            return session
        except Exception as e:
            logging.error("Error connecting to SAP: %s", e)
            raise Exception("Unable to connect to SAP GUI. Ensure SAP is running and scripting is enabled.") from e

    def execute_transaction(self, session, material, centers):
        try:
            logging.info(f"Starting transaction ME2M for material {material} and centers {centers}.")
            
            # Garantir que estamos na tela inicial
            session.findById("wnd[0]/tbar[0]/btn[3]").press()
            time.sleep(2)
            
            # Abrir a transação ME2M
            session.findById("wnd[0]").resizeWorkingPane(99, 38, 0)
            session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("F00003")
            
            # Definir o número do material
            session.findById("wnd[0]/usr/ctxtEM_MATNR-LOW").text = material
            session.findById("wnd[0]/usr/ctxtEM_MATNR-LOW").caretPosition = len(material)
            
            # Abrir a janela de seleção de centros
            session.findById("wnd[0]/usr/btn%_EM_WERKS_%_APP_%-VALU_PUSH").press()
            time.sleep(1)
            
            # Selecionar os centros
            for i, center in enumerate(centers):
                field_path = f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{i}]"
                session.findById(field_path).text = center
                session.findById(field_path).setFocus()
                session.findById(field_path).caretPosition = len(center)
            
            # Confirmar seleção
            session.findById("wnd[1]/tbar[0]/btn[8]").press()
            time.sleep(1)
            
            # Definir "Abrangência da lista"
            session.findById("wnd[0]/usr/ctxtLISTU").text = "BEST ALV"
            session.findById("wnd[0]/usr/ctxtLISTU").setFocus()
            session.findById("wnd[0]/usr/ctxtLISTU").caretPosition = len("BEST ALV")
            
            # Executar a transação
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            logging.info(f"Transaction ME2M completed for material {material} and centers {centers}.")
            time.sleep(2)
            
            # Exportar a planilha e armazenar na memória
            return self.export_spreadsheet(session, material)
            
        except Exception as e:
            logging.error("Error executing transaction: %s", e)
            return None

    def export_spreadsheet(self, session, material):
        try:
            logging.info("Exporting spreadsheet from SAP.")
            
            # Clicar no botão de exportação
            session.findById("wnd[0]/tbar[1]/btn[43]").press()
            time.sleep(2)
            
            # Selecionar formato de planilha
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            time.sleep(2)
            
            # Nome temporário para manter apenas na memória
            temp_filename = f"EXPORT_{material}.xlsx"
            
            file_field = "wnd[1]/usr/ctxtDY_FILENAME"
            try:
                file_input = session.findById(file_field)
                file_input.text = temp_filename
                file_input.caretPosition = len(temp_filename)
                time.sleep(1)
            except Exception:
                logging.error("Filename field not found. Check SAP GUI Scripting Recorder for correct ID.")
                return None
            
            # Confirmar exportação
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
            time.sleep(5)  # Espera extra para garantir que o SAP conclua a exportação
            
            logging.info(f"Spreadsheet export completed successfully: {temp_filename}")
            
            # Fechar janela de exportação
            session.findById("wnd[0]/tbar[0]/btn[3]").press()
            time.sleep(2)
            
            return temp_filename
            
        except Exception as e:
            logging.error("Error exporting spreadsheet: %s", e)
            return None


def main():
    automation = SAPAutomation()
    session = automation.connect_to_sap()
    
    # Ler lista de materiais do arquivo Excel
    materials_file = os.path.join(os.getcwd(), "lista_de_NMs.xlsx")
    materials_df = pd.read_excel(materials_file)
    materials = materials_df["Minimum Lot Size[NM]"].astype(str).tolist()
    
    centers = ["2914", "2096", "2032", "20AI", "20AF"]  # Lista de centros
    
    merged_data = []

    for material in materials:
        temp_file = automation.execute_transaction(session, material, centers)
        time.sleep(2)  # Pequena pausa para garantir que o sistema esteja pronto para o próximo material
        
        if temp_file and os.path.exists(temp_file):
            df = pl.read_excel(temp_file)
            df = df.with_columns(pl.lit(material).alias("Material"))
            merged_data.append(df)
            os.remove(temp_file)  # Remover arquivo temporário após leitura
    
    if merged_data:
        final_df = pl.concat(merged_data)
        final_path = os.path.join(os.getcwd(), "merged_data.xlsx")
        final_df.write_excel(final_path)
        logging.info(f"Merged data saved to {final_path}")
    
    input("Pressione Enter para encerrar o script e sair do SAP...")

if __name__ == "__main__":
    main()




