import win32com.client
import time
import logging

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

    def execute_transaction(self, session, material_number, center):
        try:
            logging.info("Starting transaction and entering material and center.")
            # Resize the working pane
            ID = session.findById("wnd[0]")
            ID.resizeWorkingPane(99, 38, 0)

            # Set transaction MD04
            okcd_id = session.findById("wnd[0]/tbar[0]/okcd")
            okcd_id.Text = "MD04"
            okcd_id.SetFocus()
            ID.sendVKey(0)  # Execute
            time.sleep(2)

            # Input the material number
            material_field_id = "wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR"
            material_field = session.findById(material_field_id)
            material_field.Text = material_number

            # Set caret position
            material_field.SetFocus()
            time.sleep(1)

            # Input the center
            center_field_id = "wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS"
            center_field = session.findById(center_field_id)
            center_field.Text = center

            # Set focus and caret position for center field
            center_field.SetFocus()
            time.sleep(1)

            # Execute the query
            ID.sendVKey(0)  # Execute
            time.sleep(2)

            # Set caret position for the material field in the details screen
            detail_material_field_id = "wnd[0]/usr/subINCLUDE8XX:SAPMM61R:0800/ctxtRM61R-MATNR"
            detail_material_field = session.findById(detail_material_field_id)
            detail_material_field.SetFocus()
            time.sleep(1)

            # Set caret position
            detail_material_field.CaretPosition = 6

            # Send VKey to trigger the next action
            ID.sendVKey(2)  # Adjust according to the action needed
            time.sleep(2)

            logging.info("Transaction completed successfully.")
            return "MD04 processing completed successfully."
        except Exception as e:
            logging.error("Error executing transaction for Material %s, Center %s: %s", material_number, center, e)
            return "Error"

    def extract_mrp1_data(self, session):
        """Extrai e imprime os dados específicos na aba MRP1."""
        try:
            # Tente mudar para a aba MRP1
            tab_mrp1_id = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13")  # ID da aba MRP1
            tab_mrp1_id.select()  # Seleciona a aba MRP1
            logging.info("Changed to MRP1 tab successfully.")
            time.sleep(1)  # Aguarda um pouco para garantir que a aba está totalmente carregada

            # Extrai os valores dos campos Ponto reabastec. e Estoque máximo
            ponto_reabastec = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/txtMARC-MINBE").Text
            estoque_maximo = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2483/txtMARC-MABST").Text

            # Impressão minimalista
            print(f"Ponto reabastec.: {ponto_reabastec}, Estoque máximo: {estoque_maximo}")

        except Exception as e:
            logging.error("Error extracting data from MRP1 tab: %s", e)

    def extract_mrp2_data(self, session):
        """Extrai e imprime os dados específicos na aba MRP2."""
        try:
            # Tente mudar para a aba MRP2
            tab_mrp2_id = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14")  # ID da aba MRP2
            tab_mrp2_id.select()  # Seleciona a aba MRP2
            logging.info("Changed to MRP2 tab successfully.")
            time.sleep(1)  # Aguarda um pouco para garantir que a aba está totalmente carregada

            # Extrai os valores dos campos AAP proposta e Prz. entrg. prev.
            aap_proposta = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-VSPVB").Text
            prz_entrg_prev = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-PLIFZ").Text

            # Impressão minimalista
            print(f"AAP proposta: {aap_proposta}, Prz. entrg. prev.: {prz_entrg_prev}")

        except Exception as e:
            logging.error("Error extracting data from MRP2 tab: %s", e)

    def extract_financial_data(self, session):
        """Extrai o campo Preço médio móvel na aba Contabilidade fin.1."""
        try:
            # Tente mudar para a aba Contabilidade fin.1
            tab_financial_id = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP25")  # ID da aba Contabilidade fin.1
            tab_financial_id.select()  # Seleciona a aba
            logging.info("Changed to Contabilidade fin.1 tab successfully.")
            time.sleep(1)  # Aguarda um pouco para garantir que a aba está totalmente carregada

            # Extrai o valor do campo Preço médio móvel
            preco_medio_movel = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP25/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB1:SAPLCKMMAT:0010/tabsTABS/tabpPPLF/ssubSUBML:SAPLCKMMAT:0300/txtMBEW-VERPR").Text

            print(f"Preço médio móvel: {preco_medio_movel}")

        except Exception as e:
            logging.error("Error extracting data from Contabilidade fin.1 tab: %s", e)

    def extract_all_materials(self, session, center):
        """Extrai e imprime todos os números de material disponíveis no centro especificado."""
        try:
            # Tente mudar para a aba de lista de materiais
            tab_materials_id = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14")  # ID da aba de lista de materiais
            tab_materials_id.select()  # Seleciona a aba de lista de materiais
            logging.info("Changed to materials tab successfully.")
            time.sleep(1)  # Aguarda um pouco para garantir que a aba está totalmente carregada

            # Aqui, você deve implementar a lógica específica para extrair os números de material do centro
            # Isso pode variar dependendo da estrutura da interface do SAP

            # Para fins de exemplo, vamos imprimir uma mensagem
            print(f"Extraindo números de material disponíveis no centro {center}...")
            # A lógica para listar os materiais deve ser implementada aqui

        except Exception as e:
            logging.error("Error extracting materials from center %s: %s", center, e)

# def main():
#     material_number = "11624543"  # Numero do Material
#     center = "2096"  # Centro
#     automation = SAPAutomation()

#     # Conectar ao SAP
#     session = automation.connect_to_sap()

#     # Executar a transacao
#     result = automation.execute_transaction(session, material_number, center)

#     # Extrair dados da aba MRP1
#     automation.extract_mrp1_data(session)

#     # Extrair dados da aba MRP2
#     automation.extract_mrp2_data(session)

#     # Extrair dados da aba Contabilidade fin.1
#     automation.extract_financial_data(session)

#     # Extrair todos os números de material disponíveis no centro
#     automation.extract_all_materials(session, center)

#     # Pausar a execucao para que o usuario possa ver a tela
#     input("Pressione Enter para encerrar o script e sair do SAP...")

# if __name__ == "__main__":
#     main()

def main():
    materials = ["11624543", "10001083"]  # Lista de números de materiais
    center = "2096"  # Centro
    automation = SAPAutomation()

    # Conectar ao SAP
    session = automation.connect_to_sap()

    for material_number in materials:
        # Executar a transação
        result = automation.execute_transaction(session, material_number, center)

        # Extrair dados da aba MRP1
        automation.extract_mrp1_data(session)

        # Extrair dados da aba MRP2
        automation.extract_mrp2_data(session)

        # Extrair dados da aba Contabilidade fin.1
        automation.extract_financial_data(session)

        # Voltar para a página de input de material
        # Ajuste o código abaixo conforme necessário para retornar à página correta
        ID = session.findById("wnd[0]")
        ID.resizeWorkingPane(99, 38, 0)  # Redimensionar o painel de trabalho
        okcd_id = session.findById("wnd[0]/tbar[0]/okcd")
        okcd_id.Text = "MD04"  # Definir a transação novamente
        okcd_id.SetFocus()
        ID.sendVKey(0)  # Executar
        time.sleep(2)  # Esperar a transação carregar

    # Pausar a execução para que o usuário possa ver a tela
    input("Pressione Enter para encerrar o script e sair do SAP...")

if __name__ == "__main__":
    main()


