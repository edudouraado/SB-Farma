#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sistema Completo de Vendas Canceladas - INTEGRAÇÃO RPA + EMAIL
Extração automática no Alpha7 e envio de relatórios via e-mail.
"""

import pyautogui
import pyperclip
import time
import pandas as pd
import os
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

class SistemaVendasCanceladasFinal:
    def __init__(self):
        """Inicializa o sistema (RPA, Diretórios e E-mails)"""
        
        # --- Configurações de segurança do PyAutoGUI ---
        pyautogui.PAUSE = 0.5
        pyautogui.FAILSAFE = True

        # --- Diretórios ---
        self.pasta_bruto = r"C:\Users\Eduardo\Documents\GitHub\SB-Farma\vendas_canceladas\arquivo bruto"
        self.pasta_destino = "Arquivos_Individuais_Canceladas"
        self.df = None
        self.arquivo_origem = None # Será calculado dinamicamente
        
        # --- Configurações de email (DIRETO NO CÓDIGO) ---
        # Removidos os espaços da senha para funcionar corretamente no envio
        self.config_email = {
            'smtp_server': "smtp.gmail.com",
            'smtp_port': 587,
            'email_remetente': "eduardo.dourado@sbfarma.com.br",
            'senha_remetente': "lugervwgifvsaaio"
        }
        
        # --- Emails das lojas (AMBIENTE DE TESTE) ---
        self.emails_lojas = {
            1: "eduardo.dourado@sbfarma.com.br", 2: "eduardo.dourado@sbfarma.com.br", 3: "eduardo.dourado@sbfarma.com.br", 
            # 4: "loja4@sbfarma.com.br", 5: "loja5@sbfarma.com.br", 6: "loja6@sbfarma.com.br",
            # 7: "loja7@sbfarma.com.br", 8: "loja8@sbfarma.com.br", 9: "loja9@sbfarma.com.br",
            # (Adicione ou descomente o restante depois dos testes)
        }

    # ==========================================
    # ETAPAS DO ROBÔ (PYAUTOGUI)
    # ==========================================
    def abrir_e_logar_sistema(self):
        print("🤖 Buscando Alpha7 no Windows...")
        pyautogui.press('win')
        time.sleep(1)
        pyautogui.write('A7', interval=0.1)
        time.sleep(1)
        pyautogui.press('enter')
        
        print("⏳ Aguardando tela de login (8 segundos)...")
        time.sleep(8)
        
        # Inserir Usuário
        pyautogui.click(x=1066, y=470)
        pyautogui.write('EDUARDO', interval=0.1)
        
        # Inserir Senha
        pyautogui.click(x=1063, y=499)
        pyautogui.write('SBFARMA2025', interval=0.1)
        
        # Clicar em Entrar
        pyautogui.click(x=1002, y=571)
        
        print("✅ Login efetuado com sucesso!")
        time.sleep(8)

    def navegar_para_relatorio(self):
        print("📂 Navegando para 'venda por item'...")
        pyautogui.click(x=464, y=258)
        time.sleep(1)
        pyautogui.click(x=724, y=306)
        pyautogui.write('venda por item', interval=0.1)
        pyautogui.press('enter')
        
        print("⏳ Aguardando carregamento da tela (5 segundos)...")
        time.sleep(5)
        print("✅ Tela 'venda por item' aberta.")

    def preencher_filtros_e_datas(self):
        print("⚙️ Aplicando filtros de status...")
        pyautogui.click(x=535, y=314)
        time.sleep(1)
        pyautogui.click(x=470, y=435)
        time.sleep(0.5)
        pyautogui.click(x=535, y=314)
        time.sleep(1)
        
        print("📅 Calculando e preenchendo intervalo de datas...")
        data_inicial = (datetime.now() - timedelta(days=3)).strftime('%d/%m/%Y')
        data_final = (datetime.now() - timedelta(days=1)).strftime('%d/%m/%Y')
        
        pyautogui.click(x=521, y=385)
        time.sleep(1)
        
        # Data Inicial
        pyautogui.click(x=588, y=508)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(data_inicial, interval=0.1)
        time.sleep(0.5)
        
        # Data Final
        pyautogui.click(x=555, y=536)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(data_final, interval=0.1)
        time.sleep(0.5)
        
        print(f"✅ Datas configuradas para: {data_inicial} até {data_final}")

    def configurar_layout_e_atualizar(self):
        print("📐 Ajustando o layout das colunas (Usuário e Embalagem)...")
        pyautogui.click(x=521, y=385)
        time.sleep(1)
        
        # Configurar 'usuário'
        pyautogui.click(x=706, y=491)
        time.sleep(0.5)
        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.write('usu', interval=0.1)
        time.sleep(0.5)
        
        # Configurar 'embalagem'
        pyautogui.click(x=650, y=521)
        time.sleep(1)
        pyautogui.click(x=614, y=717)
        time.sleep(0.5)
        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.write('emba', interval=0.1)
        time.sleep(0.5)
        
        # Desmarcar caixas
        print("🔳 Desmarcando opções extras...")
        pyautogui.click(x=1000, y=496)
        time.sleep(0.5)
        pyautogui.click(x=704, y=520)
        time.sleep(0.5)
        
        # Atualizar
        print("🔄 Solicitando atualização dos dados ao sistema...")
        pyautogui.click(x=1414, y=558)
        
        print("⏳ Aguardando o relatório carregar na tela (8 segundos)...")
        time.sleep(8)
        print("✅ Dados carregados!")

    def exportar_relatorio(self):
        print("💾 Iniciando processo de exportação do arquivo...")
        pyautogui.click(x=465, y=558)
        time.sleep(2) 
        
        data_ini_arq = (datetime.now() - timedelta(days=3)).strftime('%d-%m-%Y')
        data_fim_arq = (datetime.now() - timedelta(days=1)).strftime('%d-%m-%Y')
        
        nome_arquivo = f"Vendas_Canceladas_{data_ini_arq}_a_{data_fim_arq}.csv"
        caminho_completo = f"{self.pasta_bruto}\\{nome_arquivo}"
        
        print(f"📋 Salvando em: {caminho_completo}")
        
        pyperclip.copy(caminho_completo)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        
        pyautogui.click(x=877, y=787)
        time.sleep(1)
        
        pyautogui.click(x=879, y=945)
        time.sleep(1)
        
        print("⬇️ Salvando o arquivo...")
        pyautogui.click(x=1093, y=825)
        
        print("⏳ Aguardando o download ser concluído (5 segundos)...")
        time.sleep(5)
        print("✅ Relatório exportado com sucesso!")

    # ==========================================
    # ETAPAS DE PROCESSAMENTO (PANDAS E E-MAIL)
    # ==========================================
    def carregar_dados(self):
        data_ini_arq = (datetime.now() - timedelta(days=3)).strftime('%d-%m-%Y')
        data_fim_arq = (datetime.now() - timedelta(days=1)).strftime('%d-%m-%Y')
        nome_arquivo = f"Vendas_Canceladas_{data_ini_arq}_a_{data_fim_arq}.csv"
        
        self.arquivo_origem = os.path.join(self.pasta_bruto, nome_arquivo)
        
        if not os.path.exists(self.arquivo_origem):
            print(f"❌ Arquivo não encontrado: {self.arquivo_origem}")
            return False
        
        try:
            print(f"📂 Carregando arquivo extraído: {nome_arquivo}")
            self.df = pd.read_csv(self.arquivo_origem, decimal=',')
            
            print(f"📊 Dados carregados: {len(self.df)} registros")
            
            colunas_necessarias = [
                'Cód. Un. Neg.', 'Usuário', 'Cód. Barras/Etiq.', 
                'Embalagem', 'Itens', 'Venda', '% Tot.', 'Desconto', '%'
            ]
            
            colunas_faltando = [col for col in colunas_necessarias if col not in self.df.columns]
            
            if colunas_faltando:
                print("⚠️  Colunas não encontradas: " + str(colunas_faltando))
                return False
            
            self.df = self.df[colunas_necessarias]
            print(f"✅ Dados filtrados para {len(colunas_necessarias)} colunas")
            return True
            
        except Exception as e:
            print("❌ Erro ao carregar dados: " + str(e))
            return False
    
    def criar_pasta_destino(self):
        if not os.path.exists(self.pasta_destino):
            os.makedirs(self.pasta_destino)
            print(f"📁 Pasta criada: {self.pasta_destino}")
    
    def gerar_tabela_html(self, dados_loja):
        tabela = '<table border="1" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">'
        tabela += '<thead style="background-color: #007bff; color: white;"><tr>'
        tabela += '<th style="padding: 8px; text-align: left;">Cód. Barras</th>'
        tabela += '<th style="padding: 8px; text-align: left;">Produto</th>'
        tabela += '<th style="padding: 8px; text-align: left;">Embalagem</th>'
        tabela += '<th style="padding: 8px; text-align: right;">Valor</th>'
        tabela += '<th style="padding: 8px; text-align: left;">Usuário</th>'
        tabela += '</tr></thead><tbody>'
        
        for _, row in dados_loja.iterrows():
            tabela += '<tr>'
            tabela += '<td style="padding: 6px;">' + str(row['Cód. Barras/Etiq.']) + '</td>'
            qtd_itens = int(row['Itens']) if pd.notnull(row['Itens']) else 0
            tabela += '<td style="padding: 6px;">' + str(qtd_itens) + '</td>'
            tabela += '<td style="padding: 6px;">' + str(row['Embalagem']) + '</td>'
            valor_formatado = "R$ " + "{:,.2f}".format(row['Venda']).replace(',', 'X').replace('.', ',').replace('X', '.')
            tabela += '<td style="padding: 6px; text-align: right;">' + valor_formatado + '</td>'
            tabela += '<td style="padding: 6px;">' + str(row['Usuário']) + '</td>'
            tabela += '</tr>'
        
        tabela += '</tbody></table>'
        return tabela
    
    def segregar_por_loja(self):
        if self.df is None: return []
        
        self.df['Cód. Un. Neg.'] = pd.to_numeric(self.df['Cód. Un. Neg.'], errors='coerce')
        self.df = self.df.dropna(subset=['Cód. Un. Neg.'])
        self.df['Cód. Un. Neg.'] = self.df['Cód. Un. Neg.'].astype(int)
        
        lojas_unicas = sorted(self.df['Cód. Un. Neg.'].unique())
        print("🏪 Lojas encontradas: " + str(lojas_unicas))
        
        arquivos_criados = []
        for loja in lojas_unicas:
            dados_loja = self.df[self.df['Cód. Un. Neg.'] == loja].copy()
            qtd_registros = len(dados_loja)
            
            nome_arquivo = f"Vendas_Canceladas_Loja_{str(loja).zfill(2)}_{datetime.now().strftime('%Y%m%d')}.xlsx"
            caminho_completo = os.path.join(self.pasta_destino, nome_arquivo)
            
            dados_loja.to_excel(caminho_completo, index=False)
            print(f"✅ Loja {str(loja).zfill(2)}: {qtd_registros} registros → {nome_arquivo}")
            arquivos_criados.append((loja, nome_arquivo, qtd_registros))
            
        print(f"\n📁 Total de arquivos criados: {len(arquivos_criados)}")
        return arquivos_criados

    def gerar_corpo_email(self, loja, qtd_registros, valor_total, dados_loja):
        tabela_html = self.gerar_tabela_html(dados_loja)
        corpo_email = f"""
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
        .header {{ background-color: #f8f9fa; padding: 15px; border-left: 4px solid #007bff; margin-bottom: 20px; }}
        .summary {{ background-color: #e9ecef; padding: 10px; border-radius: 5px; margin-bottom: 20px; }}
        .instructions {{ background-color: #fff3cd; padding: 15px; border-left: 4px solid #ffc107; margin: 20px 0; }}
        .footer {{ margin-top: 30px; padding-top: 20px; border-top: 1px solid #dee2e6; }}
    </style>
</head>
<body>
    <div class="header">
        <h2>📊 Relatório de Vendas Canceladas - Loja {str(loja).zfill(2)}</h2>
        <p><strong>Data:</strong> {datetime.now().strftime('%d/%m/%Y')}</p>
    </div>
    <div class="summary">
        <h3>📋 Resumo</h3>
        <p><strong>Quantidade de itens cancelados:</strong> {qtd_registros}</p>
        <p><strong>Valor total:</strong> R$ {"{:,.2f}".format(valor_total).replace(',', 'X').replace('.', ',').replace('X', '.')}</p>
    </div>
    <h3>🛍️ Detalhes dos Itens Cancelados</h3>
    {tabela_html}
    <div class="instructions">
        <h3>⚠️ Ação Necessária</h3>
        <p><strong>Solicitamos que seja feita a verificação física dos itens listados acima:</strong></p>
        <ul>
            <li>✅ Verificar se os produtos estão fisicamente na loja</li>
            <li>📸 Fotografar os itens encontrados</li>
            <li>📧 Responder este email confirmando a verificação</li>
        </ul>
        <p><strong>⏰ Prazo:</strong> 24 horas a partir do recebimento deste email</p>
    </div>
    <div class="footer">
        <p>Em anexo, você encontrará o arquivo Excel com os detalhes completos para sua análise.</p>
        <p>Para dúvidas ou esclarecimentos, entre em contato conosco.</p>
        <br>
        <p><strong>Atenciosamente,<br>
        Equipe de Controle de Prevenção de Perdas</strong></p>
    </div>
</body>
</html>
"""
        return corpo_email
    
    def enviar_email(self, loja, nome_arquivo, qtd_registros):
        if self.config_email is None or loja not in self.emails_lojas:
            return False
        
        email_destino = self.emails_lojas[loja]
        
        try:
            dados_loja = self.df[self.df['Cód. Un. Neg.'] == loja].copy()
            valor_total = dados_loja['Venda'].sum()
            
            msg = MIMEMultipart()
            msg['From'] = self.config_email['email_remetente']
            msg['To'] = email_destino
            msg['Subject'] = f"Relatório de Vendas Canceladas - Loja {str(loja).zfill(2)} - {datetime.now().strftime('%d/%m/%Y')}"
            
            corpo = self.gerar_corpo_email(loja, qtd_registros, valor_total, dados_loja)
            msg.attach(MIMEText(corpo, 'html', 'utf-8'))
            
            caminho_arquivo = os.path.join(self.pasta_destino, nome_arquivo)
            if os.path.exists(caminho_arquivo):
                with open(caminho_arquivo, "rb") as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={nome_arquivo}')
                msg.attach(part)
            
            server = smtplib.SMTP(self.config_email['smtp_server'], self.config_email['smtp_port'])
            server.starttls()
            server.login(self.config_email['email_remetente'], self.config_email['senha_remetente'])
            server.sendmail(self.config_email['email_remetente'], email_destino, msg.as_string())
            server.quit()
            
            print(f"📧 Email enviado para Loja {str(loja).zfill(2)} ({email_destino})")
            return True
        except Exception as e:
            print(f"❌ Erro ao enviar email para Loja {loja}: {e}")
            return False

    # ==========================================
    # O CÉREBRO: ORQUESTRAÇÃO GERAL
    # ==========================================
    def executar_sistema_completo(self):
        print("=" * 60)
        print("🚀 INICIANDO AUTOMAÇÃO TOTAL: EXTRAÇÃO RPA + DISPARO DE E-MAIL")
        print("=" * 60)
        
        # 1. Roda o Robô
        self.abrir_e_logar_sistema()
        self.navegar_para_relatorio()
        self.preencher_filtros_e_datas()
        self.configurar_layout_e_atualizar()
        self.exportar_relatorio()
        
        # Pausa estratégica para garantir que o arquivo terminou de salvar
        print("\n⏳ Aguardando salvamento do arquivo no Windows (3 segundos)...")
        time.sleep(3)
        
        # 2. Roda a leitura de Dados
        if not self.carregar_dados():
            print("❌ Falha ao carregar dados extraídos. Encerrando o processo de e-mails.")
            return
        
        self.criar_pasta_destino()
        arquivos_criados = self.segregar_por_loja()
        
        if not arquivos_criados:
            print("❌ Nenhum dado foi processado. Encerrando.")
            return
        
        # 3. Roda os E-mails
        print("\n📧 INICIANDO ENVIO DE EMAILS")
        print("=" * 30)
        
        emails_enviados = 0
        emails_falharam = 0
        
        for loja, nome_arquivo, qtd_registros in arquivos_criados:
            if self.enviar_email(loja, nome_arquivo, qtd_registros):
                emails_enviados += 1
            else:
                emails_falharam += 1
        
        print("\n🎉 PROCESSO TOTAL CONCLUÍDO!")
        print("=" * 25)
        print(f"📁 Relatórios gerados: {len(arquivos_criados)}")
        print(f"✅ Emails enviados com sucesso: {emails_enviados}")
        print(f"❌ Emails com falha: {emails_falharam}\n")

def main():
    sistema = SistemaVendasCanceladasFinal()
    sistema.executar_sistema_completo()

if __name__ == "__main__":
    main()