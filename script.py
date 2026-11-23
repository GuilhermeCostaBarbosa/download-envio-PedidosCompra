import pandas as pd
from time import time, sleep
import logging
import sys
from tqdm import tqdm
from pywinauto import Application
import win32com.client
import os
import glob
import threading
import json

destino_po = r'C:\Users\guilherme.Barbosa\Desktop\python\envio_po_email\pedidos_compra'

df_pedidos = pd.read_excel('Follow-up Pedidos.xlsx', sheet_name='po_liberados')

# configuracao Logs
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('extacao_sap.log', mode='a', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

# validacao planilha
def validar_planilha():
    # implementar logica
    pass

def salvar_pdf_seguro(destino_pdf, nome_arquivo, timeout=45):
    inicio = time()
    while time() - inicio < timeout:
        try:
            # Conecta na janela do Windows chamada "Salvar como" ou "Save as"
            app = Application(backend="win32").connect(title_re=".*Salvar.*como.*", timeout=1)
            janela = app.window(title_re=".*Salvar.*como.*")
            
            # Garante que a janela está pronta
            janela.wait('ready', timeout=2)
            
            # Encontra a barra de endereço do arquivo (Edit control) e injeta o texto DIRETAMENTE
            arquivo_completo = fr'{destino_pdf}\{nome_arquivo}.pdf'
            janela.Edit.set_edit_text(arquivo_completo)
            sleep(1)
            
            # Clica no botão "Salvar" (geralmente é o Button1 ou Button2 no win32)
            # Para garantir, apertar simular enter na janela
            janela.send_keystrokes('{ENTER}') 
            
            return True
        except Exception as e:
            # Continua tentando até a janela aparecer
            sleep(1)
            
    return False

session = None

# Conectar SAP
def connect_sap():
    logging.info('Tentando conectar com SAP GUI...')
    try:
        SapGuiAuto = win32com.client.GetObject('SAPGUI')
    except Exception:
        logging.critical('SAP GUI não encontrado.')
        # Implementar logica para login automatizado
        raise Exception('SAP GUI não encontrado, Abra e faça o Logun')

    application = SapGuiAuto.GetScriptingEngine

    for connection in application.Connections:
        for sess in connection.Sessions:
            try:
                sess.findById('wnd[0]')
                print('Conectado ao SAP com Sucesso.')
                logging.info('Conectado ao SAP com Sucesso.')
                return sess
            except Exception:
                pass

    logging.error('Nenhuma sessão SAP ativa foi encontrada.')
    raise Exception("Nenhuma sessão SAP ativa encontrada. Abra uma sessão e faça o login.")

# Capturar Sessao SAP
def get_sap_session():
    global session
    if session is None:
        session = connect_sap()
    return session

# Capturar elementos SAP
def wait_for_element(element_id, timeout=15):
    sess = get_sap_session()
    start_time = time()

    while time() - start_time < timeout:
        try:
            element = sess.findById(element_id)
            return element
        except Exception:
            sleep(0.5)

    logging.error(f'Timeout: o elemento {element_id} não carregou após {timeout}')
    raise Exception(f'Elemento SAP não encontrado: {element_id}')
    
def extract_order(sess, df):
    logging.info('Iniciando download dos pedidos...')

    for pedido_compra in tqdm(df.index, desc='extraindo pedidos'):
        pedido = str(df.loc[pedido_compra, 'PO']).strip()
        tipo_pedido = str(df.loc[pedido_compra, 'TIPO RC']).strip()
        # gcm = str(df.loc[pedido_compra, 'Grupo de compradores']).strip()

        if tipo_pedido == 'NB':
            try:
                logging.info(f'Extraindo pedido {pedido} do tipo {tipo_pedido}...')

                wait_for_element("wnd[0]/tbar[0]/okcd").text = "ME9F"
                sess.findById("wnd[0]").sendVKey(0)
                sess.findById("wnd[0]/usr/ctxtS_EBELN-LOW").text = f"{pedido}"
                sess.findById("wnd[0]").sendVKey(0)
                sess.findById("wnd[0]/tbar[1]/btn[8]").press()
                wait_for_element("wnd[0]/usr/chk[1,5]").selected = True
                sess.findById("wnd[0]/tbar[1]/btn[16]").press()
                sess.findById("wnd[0]/tbar[0]/btn[15]").press()
                sess.findById("wnd[0]/tbar[0]/btn[15]").press()

                logging.info(f'Pedido {pedido} extraido com sucesso...')

                
                salvar_pdf_seguro(destino_pdf=destino_po, nome_arquivo=pedido)

            except Exception as e:
                logging.error(f'Erro ao processar pedido {pedido}: {e}')
                continue

        elif tipo_pedido in ['ZSER', 'ZSTQ', 'ZATI', 'ZNB', 'ZSCO']:
            try:
                logging.info(f'Extraindo pedido {pedido} do tipo {tipo_pedido}...')

                wait_for_element("wnd[0]/tbar[0]/okcd").text = "ME29N"
                sess.findById("wnd[0]").sendVKey(0)
                wait_for_element("wnd[0]/tbar[1]/btn[17]").press()
                sess.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = f"{pedido}"
                sess.findById("wnd[1]/tbar[0]/btn[0]").press()
                sess.findById("wnd[0]/tbar[1]/btn[21]").press()
                wait_for_element("wnd[0]/usr/tblSAPDV70ATC_NAST3").getAbsoluteRow(0).selected = True
                sess.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/lblDV70A-STATUSICON[0,0]").setFocus()
                sess.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/lblDV70A-STATUSICON[0,0]").caretPosition = 0
                sess.findById("wnd[0]/tbar[1]/btn[6]").press()
                wait_for_element("wnd[0]/usr/tblSAPDV70ATC_NAST3").columns.elementAt(0).width = 4
                sess.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/lblDV70A-STATUSICON[0,0]").setFocus()
                sess.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/lblDV70A-STATUSICON[0,0]").caretPosition = 0
                sess.findById("wnd[0]").sendVKey(2)
                wait_for_element("wnd[0]/usr/chkNAST-DELET").selected = True
                sess.findById("wnd[0]/usr/cmbNAST-TDOCOVER").key = "D"
                sess.findById("wnd[0]/usr/cmbNAST-TDOCOVER").setFocus()
                wait_for_element("wnd[0]/tbar[0]/btn[3]").press()
                sess.findById("wnd[0]/tbar[0]/btn[11]").press()
                sess.findById("wnd[0]/tbar[0]/btn[15]").press()

                logging.info(f'Pedido {pedido} extraido com sucesso...')
                salvar_pdf_seguro(destino_pdf=destino_po, nome_arquivo=pedido)
            except Exception as e:
                logging.error(f'Erro ao processar pedido {pedido}: {e}')
                continue

        else:
            pass



def enviar_email(df, destino_po, email_teste):
    try:
        logging.info('Verificando e-mails cadastrados')
        with open('emails_setor.json', 'r', encoding='utf-8') as json_file:
          emails_setores = json.load(json_file)

    except Exception as e:
        logging.error('Erro ao verificar e-mails cadastrados: {e}')

    logging.info('Iniciando o envio de e-mails via Outlook...')

    try:
        # Conectando o app outlook
        outlook = win32com.client.Dispatch('outlook.application')

    except Exception as e:
        logging.critical(f'Não foi possível abrir o outlook: {e}')
        return

    # convertendo data
    df['DATA DE ENTREGA'] = df['DATA DE ENTREGA'].dt.strftime('%d/%m')

    # tirando pedidos duplicados
    df_unicos = df.drop_duplicates(subset=['PO'])

    for idx, row in tqdm(df_unicos.iterrows(), total=len(df_unicos), desc='Enviando emails'):
        pedido = str(row['PO']).strip()
        setor = str(row['SETOR']).strip()
        data_entrega = row['DATA DE ENTREGA']

        # montando caminho do pedido
        caminho_pdf = os.path.join(destino_po, f'{pedido}.pdf')

        sleep(3)
        if os.path.exists(caminho_pdf):
            if setor in emails_setores:
                try:
                    mail = outlook.CreateItem(0)

                    mail.DeleteAfterSubmit = False
                    pasta_enviados = outlook.GetNamespace("MAPI").GetDefaultFolder(5)
                    mail.SaveSentMessageFolder = pasta_enviados

                    mail.SentOnBehalfOfName = "guilherme@empresa.com.br"

                    mail.To = email_teste
                    mail.CC = ';'.join(emails_setores[setor])
                    mail.Subject = f'PO {pedido}'

                    mail.Display()

                    assinatura = mail.HTMLBody

                    corpo_email = f"""
                <p>Prezados(as), bom dia!<br>
                Tudo bem?</p>

                <p>Segue o pedido de compra <b>{pedido}</b>.</p>

                <p><b><i>Por favor, seguir as seguintes orientações:</i></b></p>
                
                <ul>
                    <li><span style="background-color: yellow;"><i>Favor, anexar o número do pedido de compra no rodapé da NF e no Xped do XML.</i></span></li>
                    
                    <li><span style="background-color: yellow;"><i>NF's de serviço devem ser emitidas até o dia 20 do mês, caso contrário, deverão ser canceladas e ser emitidas no mês posterior.</i></span></li>
                    
                    <li>Encaminhar a NF nos seguintes e-mails: <br>
                    <span style="background-color: yellow;"><u>recuperacaohc@empresa.com.br</u> / 
                    <u>{emails_setores[setor][0]}</u> / 
                    <u>almoxarifado@empresa.com.br</u> 
                    <b>(caso contrário, poderá impactar no pagamento)</b></span></li>
                    
                    <li>Unidade: Hospital | CNPJ: 12.123.456/0001-78.</li>
                    
                    <li>Pagamento: faturado 30 dias.</li>
                    
                    <li>Prazo: {data_entrega}.</li>
                    
                    <li>Setor: {setor}.</li>
                    
                    <li>Recebimento: Endereço completo de entrega, das 08:00 às 16:30.</li>
                </ul>

                <p>Favor, confirmar o recebimento do pedido.</p>
                
                <p>Qualquer dúvida estou à disposição.</p>
                """

                    mail.HTMLBody = corpo_email + assinatura                    
                    mail.Attachments.Add(caminho_pdf)
                    mail.Save()
                    logging.info(f'E-mail do pedido {pedido} enviado com sucesso para {email_teste}')

                    sleep(1)
                except Exception as e:
                    logging.error(f'Erro ao enviar e-mail para o pedido {pedido}: {e}')
        else:
            logging.warning(f'O PDF do pedido {pedido} não foi encontrado na pasta. E-mail não enviado.')

    logging.info('Iniciando a limpeza da pasta de pedidos...')
    
    # selecionar todos arquivos pdf
    arquivos_pdf = glob.glob(os.path.join(destino_po, "*.pdf"))
    
    for arquivo in arquivos_pdf:
        try:
            sleep(2)
            os.remove(arquivo)
            logging.info(f'Arquivo excluído: {arquivo}')
        except Exception as e:
            logging.error(f'Erro ao excluir arquivo {arquivo}: {e}')
            
    logging.info('Processo completo finalizado com sucesso!')


sessao = get_sap_session()
extract_order(sess=sessao, df=df_pedidos)

meu_email = 'guilherme@empresa.com.br'
enviar_email(df=df_pedidos, destino_po=destino_po, email_teste=meu_email)