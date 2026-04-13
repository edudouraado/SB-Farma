import pyautogui
import time
import pyperclip
import pandas as pd
import os
import glob
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta

# ==========================================
# 1. CONFIGURAÇÕES DE ACESSO GMAIL E SISTEMA
# ==========================================
GMAIL_USER = 'vanderlei.freitas@sbfarma.com.br'  
GMAIL_APP_PASSWORD = 'fpiw hlkl fdtn llwe'

USUARIO_A7 = 'EDUARDO'
SENHA_A7 = 'SBFARMA2025'

# ==========================================
# 2. CONFIGURAÇÕES GERAIS E SEGURANÇA
# ==========================================
pyautogui.FAILSAFE = True 
pyautogui.PAUSE = 0.5 

# Caminho base
CAMINHO_PASTA_BASE = r"C:\Users\Eduardo\Documents\GitHub\SB-Farma\extracao_pbm\relatorios"

EMAILS_DAS_LOJAS = {
    "1 - PF APUIARES": "eduardo.dourado@sbfarma.com.br", 
    "2 - UP TRAIRI": "eduardo.dourado@sbfarma.com.br",
    "3 - GM MARACANAÚ": "eduardo.dourado@sbfarma.com.br",
    "4 - UP MARANGUAPE": "eduardo.dourado@sbfarma.com.br",
    "5 - UP M.LUCENA": "eduardo.dourado@sbfarma.com.br",
    "6 - PF PENTECOSTE": "eduardo.dourado@sbfarma.com.br",
    "8 - UP PARAIPABA": "",
    "7 - UP ITAPIPOCA": " ",
    "9 - UP PARACURU": "",
    "10 - UP MONTESE ": "",
    "11 - UP CANINDE LJ1 ": "",
    "12 - PF URUBURETAMA LJ1 ": "",
    "13 - PF URUBURETAMA LJ2 ": "",
    "14 - MP ITAPIPOCA": "",
    "15 - PF AMONTADA LJ1 ": "",
    "16 - UP S.G. AMARANTE ": "",
    "17 - UP ITAPAJE ": "",
    "18 - PF AMONTADA LJ2 ": "",
    "19 - UP SOBRAL LJ1 ": "",
    "22 - UP CANINDE LJ2 ": "",
    "23 - MP CANINDE ": "",
    "27 - UP L. DO NORTE ": "",
    "28 - UP JUREMA ": "",
    "29 - UP ITAITINGA ": "",
    "30 - GM PRA. FERREIRA ": "",
    "31 - UP AQUIRAZ ": "",
    "32 - PF SAO BENEDITO ": "",
    "33 - UP VICOSA DO CEARA ": "",
    "34 - UP TEJUCUOCA ": "",
    "35 - MP TRAIRI ": "",
    "36 - UP FRANCISCO SÁ ": "",
    "37 - UP IGUATU ": "",
    "38 - UP TIANGUÁ ": "",
    "39 - UP CONJ. CEARÁ ": "",
    "42 - UP RUSSAS ": "",
    "43 - UP BELA CRUZ ": "",
    "45 - FARMA VIP ITAREMA ": "",
}

def obter_saudacao():
    hora = datetime.now().hour
    if 5 <= hora < 12: return "Bom dia"
    elif 12 <= hora < 18: return "Boa tarde"
    else: return "Boa noite"

# ==========================================
# ETAPA 0: ABRIR SISTEMA E LOGAR
# ==========================================
def abrir_e_logar_alpha7():
    print("\n--- ETAPA 0: ABRIR SISTEMA E LOGAR ---")
    pyautogui.press('win')
    time.sleep(1) 
    pyautogui.write('A7', interval=0.1)
    time.sleep(1)
    pyautogui.press('enter')
    print("Aguardando o Alpha7 carregar...")
    time.sleep(10)
    
    pyautogui.write(USUARIO_A7, interval=0.1)
    pyautogui.click(x=1057, y=497)
    pyautogui.write(SENHA_A7, interval=0.1)
    pyautogui.click(x=1011, y=574)
    
    print("Sessão iniciada. Carregando ecrã inicial...")
    time.sleep(8)

# ==========================================
# ETAPA 1: EXTRAÇÃO DO SISTEMA
# ==========================================
def automatizar_extracao_pbm():
    print("\n--- ETAPA 1: EXTRAÇÃO DO SISTEMA ---")
    pyautogui.hotkey('ctrl', 'space')
    pyautogui.write('layout', interval=0.1)
    pyautogui.press('enter')
    time.sleep(1.5) 
    pyautogui.write('pbm', interval=0.1)
    pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(2) 
    
    pyautogui.click(x=1124, y=416)
    pyperclip.copy(CAMINHO_PASTA_BASE)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    time.sleep(1) 
    
    ontem = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y") 
    pyautogui.click(x=898, y=469)
    pyautogui.write(ontem, interval=0.1)
    pyautogui.click(x=1057, y=477)
    pyautogui.write(ontem, interval=0.1)
    pyautogui.click(x=1069, y=723)
    print(f"Relatório de {ontem} solicitado.")

# ==========================================
# ETAPAS 2 e 3: SEPARAÇÃO E ENVIO (COM PASTA POR DATA)
# ==========================================
def separar_e_enviar_emails():
    print("\n--- ETAPA 2 e 3: PROCESSAMENTO E ENVIO ---")
    
    # Busca o Excel original na pasta raiz
    arquivos_excel = glob.glob(os.path.join(CAMINHO_PASTA_BASE, "*.xls*"))
    if not arquivos_excel:
        print("Erro: Nenhum ficheiro Excel bruto encontrado.")
        return False
        
    arquivo_bruto = max(arquivos_excel, key=os.path.getctime)
    df = pd.read_excel(arquivo_bruto)
    
    # CORREÇÃO: Limpa os espaços dos nomes das lojas que vêm no Excel
    df['loja'] = df['loja'].astype(str).str.strip()
    lojas_unicas = df['loja'].unique()

    # CRIAÇÃO DA SUBPASTA POR DATA
    data_relatorio_str = (datetime.now() - timedelta(days=1)).strftime("%d-%m-%Y")
    pasta_do_dia = os.path.join(CAMINHO_PASTA_BASE, "Arquivos_Para_Envio", data_relatorio_str)
    os.makedirs(pasta_do_dia, exist_ok=True)
    print(f"Pasta de trabalho do dia: {data_relatorio_str}")

    saudacao = obter_saudacao()
    data_formatada = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
    except Exception as e:
        print(f"Erro Gmail: {e}")
        server = None # Modificado para não travar; se não logar, ele apenas salva os arquivos

    # CORREÇÃO: Limpa os espaços dos nomes que você digitou no Dicionário
    emails_limpos = {k.strip(): v for k, v in EMAILS_DAS_LOJAS.items()}

    for loja in lojas_unicas:
        # 1. Filtra e SALVA NA PASTA (Faz isso para TODAS as lojas encontradas)
        df_loja = df[df['loja'] == loja]
        nome_arq = str(loja).replace('/', '-').replace(':', '').strip()
        
        # CORREÇÃO: A variável estava pasta_dia, mudei para o nome correto pasta_do_dia
        caminho_anexo = os.path.join(pasta_do_dia, f"Relatorio_{nome_arq}.xlsx")
        df_loja.to_excel(caminho_anexo, index=False, engine='openpyxl')
        print(f"💾 Planilha salva para a loja: {loja}")

        # 2. Verifica se tem e-mail para enviar
        email_destino = emails_limpos.get(loja)
        if not email_destino or not email_destino.strip() or not server:
            print(f"⚠️ Envio ignorado para {loja} (Sem e-mail cadastrado ou erro no GMAIL).")
            continue

        # Resumo HTML
        resumo = df_loja.groupby('usuario_venda')['valortotal'].sum().reset_index().sort_values(by='valortotal', ascending=False)
        resumo['valortotal'] = resumo['valortotal'].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        resumo.columns = ['Colaborador', 'Total Vendido']
        tabela_html = resumo.to_html(index=False, border=1, justify='center', classes='table')

        msg = MIMEMultipart()
        msg['From'] = GMAIL_USER
        msg['To'] = email_destino.strip()
        msg['Subject'] = f"Resumo de KPIs - {data_formatada} - {loja}"

        corpo_email = f"""
        <html>
            <head>
                <style>
                    .table {{border-collapse: collapse; width: 100%; max-width: 400px; font-family: Arial, sans-serif;}}
                    .table th {{background-color: #f2f2f2; padding: 8px; text-align: left;}}
                    .table td {{padding: 8px; border-bottom: 1px solid #ddd;}}
                </style>
            </head>
            <body>
                <p>{saudacao}, gerente da <b>{loja}</b>,</p>
                <p>Resumo de performance PBM de ontem (<b>{data_formatada}</b>):</p>
                {tabela_html}
                <p><br>O relatório detalhado segue em anexo.</p>
                <p>Atenciosamente,<br><b>Vanderlei Freitas</b><br>Coodenador Geral de Vendas - SB Farma</p>
            </body>
        </html>
        """
        msg.attach(MIMEText(corpo_email, 'html'))

        with open(caminho_anexo, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(caminho_anexo)}')
            msg.attach(part)

        server.send_message(msg)
        print(f"✅ E-mail enviado com sucesso para: {loja}")

    if server:
        server.quit()
    return True

def executar_robo_completo():
    print("==========================================")
    print(" INICIANDO ROBÔ PBM - SB FARMA")
    print("==========================================")
    abrir_e_logar_alpha7()
    automatizar_extracao_pbm()
    time.sleep(8)
    separar_e_enviar_emails()
    print("\n==========================================")
    print(" PROCESSO 100% CONCLUÍDO!")
    print("==========================================")

if __name__ == "__main__":
    executar_robo_completo()