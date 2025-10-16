import pandas as pd
import os
import numpy as np
import mysql.connector
from sqlalchemy import create_engine
import datetime

# ==============================================================================
# 🚨 PASSO 1: CONFIGURAÇÃO FIXA E DE BANCO DE DADOS (AJUSTE OBRIGATÓRIO!)
# ==============================================================================

DELIMITADOR = ';' 
CODIFICACAO = 'latin-1' 

# Configurações de Caminho e Saída
PASTA_RELATORIOS = "Relatórios de Inventário" # Pasta para salvar os arquivos Excel

# DADOS DE CONEXÃO COM O BANCO DE DADOS (AJUSTE OBRIGATÓRIO SE FOR USAR O MYSQL!)
DB_USER = 'root'          # Ex: root
DB_PASSWORD = '12345'        # Sua senha do MySQL
DB_HOST = '127.0.0.1'            # Endereço do seu servidor MySQL
DB_DATABASE = 'inventario_db'    # Nome do banco de dados (precisa existir)
NOME_TABELA_MYSQL = 'inventario_mudancas' # Nome da tabela que o Python irá criar/alimentar

# Colunas Chave (NOMES INTERNOS USADOS APÓS O SANEAMENTO)
COLUNA_ID = 'Cod Int'         
COLUNA_DESCRICAO = 'Descrição'
COLUNA_TEORICO = 'Soma de estoque_teorico'
COLUNA_CONTADO = 'Soma de estoque_contado'
COLUNA_DIF = 'Dif'            
COLUNA_TOTAL_VALOR = 'Total_Valor_Col' # Nome interno para a coluna de valor
COLUNA_STATUS_FINAL = 'Status_Inventario' 

# ==============================================================================
# FIM DA CONFIGURAÇÃO
# ==============================================================================


def exportar_para_mysql(df, nome_loja):
    """Exporta o DataFrame de mudanças para uma tabela no MySQL."""
    print("\nIniciando exportação para MySQL...")
    if df.empty:
        print("[MYSQL] DataFrame vazio. Nada a exportar.")
        return

    try:
        # 1. Cria a URL de conexão
        db_url = f"mysql+mysqlconnector://{DB_USER}:{DB_PASSWORD}@{DB_HOST}/{DB_DATABASE}"
        engine = create_engine(db_url)

        # 2. Prepara o DataFrame para exportação
        df_export = df.copy()
        df_export['Loja'] = nome_loja
        df_export['Data_Comparacao'] = datetime.date.today()
        
        # 3. Exporta para o MySQL (Adiciona as novas linhas à tabela existente)
        df_export.to_sql(
            name=NOME_TABELA_MYSQL,
            con=engine,
            if_exists='append', 
            index=True,           
            index_label=COLUNA_ID
        )

        print(f"[SUCESSO MYSQL] {len(df_export)} registros da Loja '{nome_loja}' adicionados à tabela '{NOME_TABELA_MYSQL}' no MySQL.")
        
    except Exception as e:
        print("\n--- ERRO ao Exportar para MySQL ---")
        print("Verifique se o MySQL está rodando, o nome do DB e as credenciais.")
        print(f"Detalhes: {e}")


def carregar_e_preparar_dados(caminho_arquivo):
    """
    Carrega o arquivo CSV, saneia cabeçalhos, corrige tipos de dados e cria a coluna de Status.
    """
    try:
        print(f"\nCarregando e preparando {os.path.basename(caminho_arquivo)}...")
        
        df = pd.read_csv(caminho_arquivo, sep=DELIMITADOR, encoding=CODIFICACAO)
        
        # 1. SANEAMENTO DE CABEÇALHOS (Limpeza de espaços e caracteres invisíveis)
        df.columns = df.columns.str.replace('\xa0', ' ').str.strip() 
        df.rename(columns={
            'Cod Int': COLUNA_ID,
            'Descrição': COLUNA_DESCRICAO,
            'Soma de estoque_teorico': COLUNA_TEORICO,
            'Soma de estoque_contado': COLUNA_CONTADO,
            'Dif': COLUNA_DIF,
            'Total': COLUNA_TOTAL_VALOR 
        }, inplace=True, errors='ignore')

        # 2. Verificação das Colunas
        colunas_requeridas = [
            COLUNA_ID, COLUNA_DESCRICAO, COLUNA_TEORICO, COLUNA_CONTADO, 
            COLUNA_DIF, COLUNA_TOTAL_VALOR
        ]
        
        if not all(col in df.columns for col in colunas_requeridas):
            print(f"\n--- ERRO: Colunas Ausentes ---")
            print(f"Colunas esperadas (após limpeza): {colunas_requeridas}")
            return None
        
        # 3. CONVERSÃO DE TIPO CRÍTICA: Garante que ID e DIF sejam números inteiros.
        df[COLUNA_ID] = pd.to_numeric(df[COLUNA_ID], errors='coerce').fillna(-1).astype(int)
        df[COLUNA_DIF] = pd.to_numeric(df[COLUNA_DIF], errors='coerce').fillna(0).astype(int)

        # 4. Criação da Coluna de Status
        df[COLUNA_STATUS_FINAL] = np.select(
            condlist=[
                df[COLUNA_DIF] < 0, 
                df[COLUNA_DIF] > 0, 
                df[COLUNA_DIF] == 0 
            ],
            choicelist=[
                'PERDA',                 
                'SOBRA',                 
                'Estoque_OK'
            ],
            default='ERRO_CALCULO'
        )

        # 5. Define a coluna de ID como índice
        df.set_index(COLUNA_ID, inplace=True)
        
        return df[[
            COLUNA_DESCRICAO, COLUNA_STATUS_FINAL, COLUNA_TEORICO, 
            COLUNA_CONTADO, COLUNA_DIF, COLUNA_TOTAL_VALOR
        ]] 
        
    except FileNotFoundError:
        print(f"\n--- ERRO: Arquivo Não Encontrado ---")
        return None
    except Exception as e:
        print(f"\n--- ERRO de Leitura ---")
        print(f"Detalhes: {e}")
        return None

def comparar_inventarios(df_antes, df_depois):
    """Compara os DataFrames e identifica as mudanças."""
    print("\nIniciando comparação de inventários...")
    
    # 1. Combina os DataFrames (Full Outer Join)
    df_combinado = pd.merge(
        df_antes, 
        df_depois, 
        left_index=True, 
        right_index=True, 
        how='outer',
        suffixes=('_Antes', '_Depois')
    )
    
    coluna_status_antes = f'{COLUNA_STATUS_FINAL}_Antes'
    coluna_status_depois = f'{COLUNA_STATUS_FINAL}_Depois'
    
    MARCADOR_AUSENTE = '[AUSENTE_NO_INVENTARIO]'
    df_relatorio = df_combinado.copy()
    
    # CORREÇÃO: Resolve o FutureWarning (substitui o fillna(inplace=True))
    df_relatorio[coluna_status_antes] = df_relatorio[coluna_status_antes].fillna(MARCADOR_AUSENTE)
    df_relatorio[coluna_status_depois] = df_relatorio[coluna_status_depois].fillna(MARCADOR_AUSENTE)

    # Filtra apenas os itens que tiveram alguma mudança (o que for OK/OK é removido)
    df_relatorio['Houve_Mudanca'] = df_relatorio[coluna_status_antes] != df_relatorio[coluna_status_depois]
    df_mudancas = df_relatorio[df_relatorio['Houve_Mudanca']].copy()
    
    # 3. Classifica o tipo exato da mudança
    def classificar_mudanca(row):
        status_a = row[coluna_status_antes]
        status_d = row[coluna_status_depois]
        
        if status_a == MARCADOR_AUSENTE and status_d != MARCADOR_AUSENTE:
            return 'PRODUTO_NOVO' 
        elif status_a != MARCADOR_AUSENTE and status_d == MARCADOR_AUSENTE:
            return 'PRODUTO_REMOVIDO'
        elif status_a != status_d:
            return 'MUDANCA_DE_STATUS'
        else:
            return 'ERRO' 

    df_mudancas['Tipo_de_Mudanca'] = df_mudancas.apply(classificar_mudanca, axis=1)

    # 4. Limpeza e Seleção Final de Colunas
    df_mudancas.replace(MARCADOR_AUSENTE, pd.NA, inplace=True)
    
    df_mudancas[COLUNA_DESCRICAO] = df_mudancas[f'{COLUNA_DESCRICAO}_Antes'].fillna(df_mudancas[f'{COLUNA_DESCRICAO}_Depois'])
    
    # Renomeia as colunas para o relatório final
    df_relatorio_final = df_mudancas.rename(columns={
        f'{COLUNA_TEORICO}_Antes': 'Teorico_Antes',
        f'{COLUNA_CONTADO}_Antes': 'Contado_Antes',
        f'{COLUNA_DIF}_Antes': 'Dif_Antes',
        f'{COLUNA_TOTAL_VALOR}_Antes': 'Total_Valor_Antes', 
        f'{COLUNA_TEORICO}_Depois': 'Teorico_Depois',
        f'{COLUNA_CONTADO}_Depois': 'Contado_Depois',
        f'{COLUNA_DIF}_Depois': 'Dif_Depois',
        f'{COLUNA_TOTAL_VALOR}_Depois': 'Total_Valor_Depois', 
        f'{COLUNA_STATUS_FINAL}_Antes': 'Status_Antes',
        f'{COLUNA_STATUS_FINAL}_Depois': 'Status_Depois',
    })[['Descrição', 'Tipo_de_Mudanca', 'Status_Antes', 'Status_Depois', 
        'Teorico_Antes', 'Contado_Antes', 'Dif_Antes', 'Total_Valor_Antes',
        'Teorico_Depois', 'Contado_Depois', 'Dif_Depois', 'Total_Valor_Depois']]

    print(f"Total de Produtos Comparados: {len(df_combinado)}")
    print(f"Total de Mudanças (Status, Adição ou Remoção) Identificadas: {len(df_mudancas)}")
    
    return df_relatorio_final

def gerar_relatorio(df_relatorio, nome_loja):
    """Salva o DataFrame de mudanças em um novo arquivo Excel DENTRO DA PASTA DE RELATÓRIOS."""
    if df_relatorio.empty:
        print("\nNenhuma mudança de status significativa foi encontrada. Nenhum relatório gerado.")
        return

    # 1. Cria a pasta se ela não existir
    if not os.path.exists(PASTA_RELATORIOS):
        os.makedirs(PASTA_RELATORIOS)
        print(f"[INFO] Pasta '{PASTA_RELATORIOS}' criada no diretório atual.")

    # 2. Cria o caminho completo do arquivo
    NOME_RELATORIO = f'relatorio_mudancas_{nome_loja}.xlsx'
    CAMINHO_COMPLETO_RELATORIO = os.path.join(PASTA_RELATORIOS, NOME_RELATORIO)
    
    try:
        df_relatorio.to_excel(CAMINHO_COMPLETO_RELATORIO, engine='openpyxl', index=True, index_label=COLUNA_ID)
        print(f"\n--- SUCESSO ---")
        print(f"Relatório de mudanças da loja '{nome_loja}' gerado com sucesso em:")
        print(os.path.abspath(CAMINHO_COMPLETO_RELATORIO))
    except Exception as e:
        print(f"\n--- ERRO ao salvar o relatório ---")
        print(f"Detalhes: {e}")

def main():
    """Função principal que executa a automação para uma única loja."""
    
    print("=====================================================")
    print("         INÍCIO DO COMPARADOR DE INVENTÁRIOS         ")
    print("=====================================================")
    
    DIRETORIO_ATUAL = os.getcwd()
    print(f"DIRETÓRIO ATUAL: {DIRETORIO_ATUAL}")
    print("Certifique-se de que os arquivos CSV estão neste local.")
    
    try:
        # 1. Solicita os nomes dos arquivos ao usuário
        print("\n--- Nova Comparação ---")
        
        arquivo_antes_nome = input("-> Digite o NOME COMPLETO do inventário ANTIGO (ex: Inventário - Loja 10 - Jan.csv): ")
        arquivo_depois_nome = input("-> Digite o NOME COMPLETO do inventário NOVO (ex: Inventário - Loja 10 - Abr.csv): ")

        caminho_completo_antes = os.path.join(DIRETORIO_ATUAL, arquivo_antes_nome)
        caminho_completo_depois = os.path.join(DIRETORIO_ATUAL, arquivo_depois_nome)

        print(f"\n[DEBUG] Procurando ANTIGO em: {caminho_completo_antes}")
        print(f"[DEBUG] Procurando NOVO em: {caminho_completo_depois}")
        
        nome_loja_base = os.path.splitext(arquivo_depois_nome)[0].replace('Inventário - ', '').strip()
        if not nome_loja_base:
             nome_loja_base = 'Inventarios'

        # 2. Carrega e Prepara os Dados
        df_antes = carregar_e_preparar_dados(arquivo_antes_nome)
        df_depois = carregar_e_preparar_dados(arquivo_depois_nome)
        
        if df_antes is None or df_depois is None:
            print("\nProcesso interrompido devido a erro de leitura.")
            return

        # 3. Compara, Gera Relatório
        df_mudancas = comparar_inventarios(df_antes, df_depois)
        
        if df_mudancas.empty:
            print(f"\n[INFO] Nenhuma mudança detectada na Loja '{nome_loja_base}'. Nada a exportar.")
        
        else:
            # Opção 1: Gerar o arquivo Excel
            gerar_relatorio(df_mudancas, nome_loja_base) 

            # Opção 2: Exportar para o MySQL Workbench/Servidor
            exportar_para_mysql(df_mudancas, nome_loja_base) 

        print("\n=====================================================")
        print("Processo de comparação finalizado.")
        print("Para comparar outra loja, execute o script novamente.")
        print("=====================================================")

    except KeyboardInterrupt:
        print("\n\nProcesso interrompido pelo usuário.")
    except Exception as e:
        print(f"\nOcorreu um erro inesperado: {e}")


if __name__ == "__main__":
    main()