import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
from datetime import datetime

# --- 1. CONFIGURAÇÕES GERAIS ---
MEU_EMAIL = "eduardo.dourado@sbfarma.com.br"
MINHA_SENHA = "hrym pqie zgkp pofa"  # Sua Senha de App
SERVIDOR_SMTP = "smtp.gmail.com"
PORTA_SMTP = 587

# --- MAPEAMENTO DE EMAILS (CHAVE = ID DA LOJA) ---
# Configurei com o padrão lojaX@sbfarma.com.br conforme seu exemplo anterior.
# Se alguma loja tiver um email diferente (ex: gerente.lojaX), basta alterar a linha específica.
emails_lojas = {
    # Teste
    "Ruslan": "ruslan.r.c@brunofarma.com",

    # Lojas (Chave é o NÚMERO INTEIRO do Código)
    1: "loja1@sbfarma.com.br",   # PF APUIARES
    2: "loja2@sbfarma.com.br",   # UP TRAIRI
    3: "loja3@sbfarma.com.br",   # GM MARACANAÚ
    4: "loja4@sbfarma.com.br",   # UP MARANGUAPE
    5: "loja5@sbfarma.com.br",   # UP M.LUCENA
    6: "loja6@sbfarma.com.br",   # PF PENTECOSTE
    7: "loja7@sbfarma.com.br",   # UP ITAPIPOCA
    8: "loja8@sbfarma.com.br",   # UP PARAIPABA
    9: "loja9@sbfarma.com.br",   # UP PARACURU
    10: "loja10@sbfarma.com.br", # UP MONTESE
    11: "loja11@sbfarma.com.br", # UP CANINDE LJ1
    12: "loja12@sbfarma.com.br", # PF URUBURETAMA LJ1
    13: "loja13@sbfarma.com.br", # PF URUBURETAMA LJ2
    14: "loja14@sbfarma.com.br", # MP ITAPIPOCA
    15: "loja15@sbfarma.com.br", # PF AMONTADA LJ1
    16: "loja16@sbfarma.com.br", # UP S.G. AMARANTE
    17: "loja17@sbfarma.com.br", # UP ITAPAJE
    18: "loja18@sbfarma.com.br", # PF AMONTADA LJ2
    19: "loja19@sbfarma.com.br", # UP SOBRAL LJ1
    22: "loja22@sbfarma.com.br", # UP CANINDE LJ2
    23: "loja23@sbfarma.com.br", # MP CANINDE
    27: "loja27@sbfarma.com.br", # UP L. DO NORTE
    28: "loja28@sbfarma.com.br", # UP JUREMA
    29: "loja29@sbfarma.com.br", # UP ITAITINGA
    30: "loja30@sbfarma.com.br", # GM PRA. FERREIRA
    31: "loja31@sbfarma.com.br", # UP AQUIRAZ
    32: "loja32@sbfarma.com.br", # PF SAO BENEDITO
    33: "loja33@sbfarma.com.br", # UP VICOSA DO CEARA
    34: "loja34@sbfarma.com.br", # UP TEJUCUOCA
    35: "loja35@sbfarma.com.br", # MP TRAIRI
    36: "loja36@sbfarma.com.br", # UP FRANCISCO SÁ
    37: "loja37@sbfarma.com.br", # UP IGUATU
    38: "loja38@sbfarma.com.br", # UP TIANGUÁ
    39: "loja39@sbfarma.com.br", # UP CONJ. CEARÁ
    42: "loja42@sbfarma.com.br", # UP RUSSAS
    45: "loja45@sbfarma.com.br", # FARMA VIP ITAREMA
    46: "loja46@sbfarma.com.br", # UP ACARAU
}

# --- DEFINIÇÃO DE COLUNAS ---
COLUNAS_DESEJADAS = [
    'Código da Un. Neg.', 
    'Embalagem', 
    'Cód. Barras/Etiqueta', 
    'Lote', 
    'Data Validade',
    'Saldo'
]

# --- CÁLCULO DO MÊS DE REFERÊNCIA ---
hoje = datetime.now()
mes_seguinte = hoje.month + 1
ano_referencia = hoje.year

if mes_seguinte > 12:
    mes_seguinte = 1
    ano_referencia += 1

meses_pt = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril', 5: 'Maio', 6: 'Junho',
    7: 'Julho', 8: 'Agosto', 9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
}
texto_referencia = f"{meses_pt[mes_seguinte]}/{ano_referencia}"

# --- FUNÇÃO VISUAL (TABELA) ---
def criar_tabela_html(df):
    html_table = df.to_html(index=False, border=0, justify="left")
    
    html_table = html_table.replace(
        '<table', 
        '<table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 12px; color: #333;"'
    )
    html_table = html_table.replace(
        '<th>', 
        '<th style="background-color: #002060; color: white; border: 1px solid #b3c2d1; padding: 10px; text-align: left;">'
    )
    html_table = html_table.replace(
        '<td>', 
        '<td style="border: 1px solid #b3c2d1; padding: 8px;">'
    )
    return html_table

# --- CARREGAR DADOS ---
arquivo_base = "testee.xlsx" 

try:
    df_completo = pd.read_excel(arquivo_base)
    
    # 1. Garante que o CÓDIGO seja lido como NÚMERO (remove erros de conversão)
    if 'Código da Un. Neg.' in df_completo.columns:
        df_completo['Código da Un. Neg.'] = pd.to_numeric(df_completo['Código da Un. Neg.'], errors='coerce').fillna(0).astype(int)

    # 2. Formata a data
    if 'Data Validade' in df_completo.columns:
        df_completo['Data Validade'] = pd.to_datetime(df_completo['Data Validade'], errors='coerce').dt.strftime('%d/%m/%Y')

    # 3. Verifica colunas
    colunas_existentes = [c for c in COLUNAS_DESEJADAS if c in df_completo.columns]
    
except Exception as e:
    print(f"Erro ao ler arquivo: {e}")
    exit()

# --- LOOP DE ENVIO E ORGANIZAÇÃO ---
# Agora buscamos os CÓDIGOS únicos
ids_unicos = df_completo['Código da Un. Neg.'].unique()

print(f"Iniciando processo para {len(ids_unicos)} códigos de loja encontrados...")
print(f"Referência considerada: {texto_referencia}")

for id_loja in ids_unicos:
    
    # Verifica se o ID está no dicionário
    if id_loja in emails_lojas:
        email_destino = emails_lojas[id_loja]
        
        # Filtra pelo CÓDIGO da loja
        filtro_loja = df_completo['Código da Un. Neg.'] == id_loja
        
        # Pega as colunas desejadas
        df_final = df_completo.loc[filtro_loja, colunas_existentes].copy()

        # --- RECUPERA O NOME DA LOJA PARA USAR NO TEXTO/PASTA ---
        try:
            # Pega o primeiro nome encontrado para este código na planilha original
            nome_loja_visual = df_completo.loc[filtro_loja, 'Nome da Un. Neg.'].iloc[0]
            nome_loja_visual = str(nome_loja_visual).strip()
        except:
            nome_loja_visual = f"Loja {id_loja}" 

        # --- ORGANIZAÇÃO DE PASTAS ---
        nome_pasta_loja = nome_loja_visual.replace('/', '-').strip()
        caminho_pasta = os.path.join("Relatorios_Enviados", nome_pasta_loja)
        
        if not os.path.exists(caminho_pasta):
            os.makedirs(caminho_pasta)
            
        data_atual_str = hoje.strftime("%Y-%m-%d")
        nome_arquivo = f"Vencimentos_{data_atual_str}.xlsx"
        caminho_completo = os.path.join(caminho_pasta, nome_arquivo)
        
        df_final.to_excel(caminho_completo, index=False)
        print(f"📁 Arquivo gerado para: {nome_loja_visual} (ID: {id_loja})")

        # --- PREPARAÇÃO DO E-MAIL ---
        tabela_html = criar_tabela_html(df_final)
        
        msg = MIMEMultipart()
        msg['From'] = MEU_EMAIL
        msg['To'] = email_destino
        msg['Subject'] = f"RELATORIO DE VENCIMENTO: Preparacao para Baixa - {nome_loja_visual}"

        corpo_email = f"""
        <html>
        <body style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">
            
            <h3 style="color: #002060; border-bottom: 2px solid #002060; padding-bottom: 5px;">
                Relatório de Vencimento: {texto_referencia}
            </h3>

            <p>Prezados,</p>
            
            <p>Segue abaixo a relação de produtos identificados no sistema (aba <strong>Item Pré-Vencido</strong>) com vencimento programado para <strong>{texto_referencia}</strong>.</p>
            
            <p>O objetivo deste comunicado é antecipar a organização do estoque para a futura rotina de baixa. Solicitamos que a equipe realize os seguintes procedimentos operacionais:</p>
            
            <ol>
                <li><strong>Conferência Física:</strong> Validar se o lote e a quantidade física correspondem ao relatório;</li>
                <li><strong>Segregação:</strong> Retirar imediatamente os itens da área de venda para evitar comercialização indevida;</li>
                <li><strong>Preparação:</strong> Deixar os produtos separados e identificados para a próxima baixa.</li>
            </ol>
            
            <br>
            {tabela_html}
            <br>
            
            <p>Ressaltamos a importância desta triagem para evitar divergências de estoque. O arquivo completo segue em anexo para impressão.</p>
            
            <hr style="border: 0; border-top: 1px solid #eee;">
            <p style="font-size: 14px; color: #555;">
            Atenciosamente,<br>
            <strong>Eduardo Dourado</strong><br>
            Setor de Prevenção de Perdas
            </p>
        </body>
        </html>
        """
        
        msg.attach(MIMEText(corpo_email, 'html'))

        # Anexa o arquivo
        with open(caminho_completo, "rb") as f:
            part = MIMEApplication(f.read(), Name=nome_arquivo)
            part['Content-Disposition'] = f'attachment; filename="{nome_arquivo}"'
            msg.attach(part)

        # Envia
        try:
            server = smtplib.SMTP(SERVIDOR_SMTP, PORTA_SMTP)
            server.starttls()
            server.login(MEU_EMAIL, MINHA_SENHA)
            server.sendmail(MEU_EMAIL, email_destino, msg.as_string())
            server.quit()
            print(f"✅ E-mail enviado para: {nome_loja_visual}")
        except Exception as e:
            print(f"❌ Erro ao enviar para {nome_loja_visual}: {e}")

    else:
        # Ignora ID 0 ou nulo, mas avisa se achar um ID válido sem email
        if id_loja > 0:
            print(f"⚠️ Código '{id_loja}' encontrado na planilha mas SEM e-mail cadastrado no script.")

print("Processo finalizado com sucesso.")