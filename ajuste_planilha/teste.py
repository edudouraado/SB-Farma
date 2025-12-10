import pandas as pd
import numpy as np

# --- CONFIGURAÇÃO ---
arquivo_entrada = 'Base cadastral atualizada.csv'
arquivo_saida = 'produtos_completo_mysql.csv'

print("Lendo o arquivo...")

# Tenta ler com encoding latin1 ou utf-8
try:
    df = pd.read_csv(arquivo_entrada, sep=';', encoding='latin1', dtype=str)
except UnicodeDecodeError:
    df = pd.read_csv(arquivo_entrada, sep=';', encoding='utf-8', dtype=str)

# --- PROCESSAMENTO ---

# Selecionamos as colunas que queremos manter
colunas_principais = ['codigo_int', 'codigobarras', 'descricao', 'classificacao_geral']
colunas_filha =      ['codigo_int', 'embalagem_filha', 'descricao', 'classificacao_geral']

# 1. Cria DataFrame baseado no Código de Barras Principal
df_principal = df[colunas_principais].copy()
df_principal.columns = ['codigo_int', 'codigo_barras', 'nome', 'classificacao_geral']

# 2. Cria DataFrame baseado na Embalagem Filha
df_filha = df[colunas_filha].copy()
df_filha.columns = ['codigo_int', 'codigo_barras', 'nome', 'classificacao_geral']

# 3. Junta tudo numa lista só
df_final = pd.concat([df_principal, df_filha])

# --- LIMPEZA ---

# Função para limpar o código (remove pontos, .0, espaços, notação científica)
def limpar_codigo(codigo):
    if pd.isna(codigo) or codigo == '' or str(codigo).lower() == 'nan':
        return None
    try:
        # Remove espaços
        cod_str = str(codigo).strip()
        # Converte float para int para str (ex: "789.0" -> 789 -> "789")
        # Troca vírgula por ponto antes da conversão
        cod_limpo = str(int(float(cod_str.replace(',', '.'))))
        return cod_limpo
    except:
        return None 

print("Limpando códigos de barras...")
df_final['codigo_barras'] = df_final['codigo_barras'].apply(limpar_codigo)

# Remove quem ficou sem código de barras válido
df_final = df_final.dropna(subset=['codigo_barras'])

# Remove duplicatas (mantém o primeiro que encontrar)
df_final = df_final.drop_duplicates(subset=['codigo_barras'])

# Garante a ordem das colunas pedida
df_final = df_final[['codigo_int', 'codigo_barras', 'nome', 'classificacao_geral']]

# --- SALVAR ---
print(f"Gerando arquivo final com {len(df_final)} produtos...")
df_final.to_csv(arquivo_saida, index=False, encoding='utf-8')

print(f"SUCESSO! Arquivo '{arquivo_saida}' gerado.")