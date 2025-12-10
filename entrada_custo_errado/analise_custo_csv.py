import pandas as pd
import os
import glob

# ==============================================================================
# CONFIGURAÇÕES
# ==============================================================================
NOME_BASE_FIXA = 'Base cadastral atualizada.csv'
NOME_RELATORIO = 'Relatorio_Familias_Completo.xlsx'

COLUNAS_DESEJADAS = [
    'NUMERO_LOJA', 
    'CODIGOBARRAS', 
    'DESCRICAO_EMBALAGEM', 
    'LABORATORIO', 
    'QUANT_EMBALAGEM', 
    'CUSTO_UNIT_R$', 
    'CUSTO_FINAL_R$', 
    'CUSTO_MAX_R$', 
    'DIF_%_MAX_MIN', 
    'FORNECEDOR'
]

# ==============================================================================
# FUNÇÕES DE LIMPEZA
# ==============================================================================
def limpar_ean(serie):
    """Remove pontuação e garante texto para comparação exata"""
    return serie.astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

def carregar_csv_texto(arquivo):
    """Lê CSV mantendo zeros à esquerda e casas decimais exatas"""
    try:
        df = pd.read_csv(arquivo, sep=',', encoding='utf-8', dtype=str)
    except:
        try:
            df = pd.read_csv(arquivo, sep=';', encoding='latin1', dtype=str)
        except:
            df = pd.read_csv(arquivo, sep=';', encoding='utf-8', dtype=str)
    return df

# ==============================================================================
# 1. MAPEAMENTO DAS FAMÍLIAS (BASE CADASTRAL)
# ==============================================================================
print("\n--- RASTREADOR DE FAMÍLIAS (MÃES E FILHOS) ---")
print(f"[1/4] Mapeando Base Cadastral ({NOME_BASE_FIXA})...")

if not os.path.exists(NOME_BASE_FIXA):
    print(f"ERRO: {NOME_BASE_FIXA} não encontrado.")
    exit()

df_base = carregar_csv_texto(NOME_BASE_FIXA)
df_base.columns = df_base.columns.str.strip().str.lower()

if 'codigobarras' not in df_base.columns or 'embalagem_filha' not in df_base.columns:
    print("ERRO NA BASE: Colunas obrigatórias não encontradas.")
    exit()

# Filtra apenas linhas que definem uma relação (tem filho preenchido)
df_relacoes = df_base[df_base['embalagem_filha'].notna() & (df_base['embalagem_filha'] != '')].copy()

# Cria dicionários de busca rápida
# EAN_MAE -> EAN_FILHO
mapa_mae_para_filho = dict(zip(limpar_ean(df_relacoes['codigobarras']), limpar_ean(df_relacoes['embalagem_filha'])))
# EAN_FILHO -> EAN_MAE (Inverso, para saber quem é a mãe de um filho perdido)
mapa_filho_para_mae = dict(zip(limpar_ean(df_relacoes['embalagem_filha']), limpar_ean(df_relacoes['codigobarras'])))

print(f"--> Base Mapeada: {len(mapa_mae_para_filho)} relações encontradas.")

# ==============================================================================
# 2. LEITURA DO ARQUIVO DE DADOS
# ==============================================================================
print("\n[2/4] Selecione o arquivo de dados:")
csvs = [f for f in glob.glob("*.csv") if NOME_BASE_FIXA not in f and "Relatorio" not in f]

if not csvs:
    print("Nenhum arquivo CSV de dados encontrado.")
    exit()

for i, f in enumerate(csvs):
    print(f"[{i+1}] {f}")

try:
    idx = int(input("Digite o número: ")) - 1
    arquivo_banco = csvs[idx]
except:
    arquivo_banco = csvs[0]

print(f"--> Lendo: {arquivo_banco}...")
df_banco = carregar_csv_texto(arquivo_banco)
df_banco.columns = df_banco.columns.str.strip().str.upper()

if 'CODIGOBARRAS' not in df_banco.columns:
    print("ERRO: Coluna CODIGOBARRAS ausente no arquivo selecionado.")
    exit()

df_banco['EAN_LIMPO'] = limpar_ean(df_banco['CODIGOBARRAS'])

# ==============================================================================
# 3. PROCESSAMENTO (AGRUPANDO POR FAMÍLIA)
# ==============================================================================
print("\n[3/4] Identificando Mães e Filhos (presentes ou isolados)...")

# Dicionário para agrupar: Chave = EAN DA MÃE (Mesmo que só tenhamos o filho, usaremos o EAN da mãe como ID do grupo)
# Estrutura: familias[ean_mae] = { 'maes': [rows], 'filhos': [rows] }
familias = {}

total_produtos_analisados = 0

for idx, row in df_banco.iterrows():
    ean_atual = row['EAN_LIMPO']
    eh_mae = ean_atual in mapa_mae_para_filho
    eh_filho = ean_atual in mapa_filho_para_mae
    
    # Se não for nem mãe nem filho (produto avulso sem relação na base), ignoramos conforme seu pedido
    if not eh_mae and not eh_filho:
        continue
        
    total_produtos_analisados += 1
    
    # Prepara os dados da linha
    dados_row = {k: v for k, v in row.to_dict().items() if k in COLUNAS_DESEJADAS}
    
    # CASO 1: É UMA MÃE
    if eh_mae:
        ean_familia = ean_atual # O ID da família é o próprio EAN desta mãe
        
        if ean_familia not in familias:
            familias[ean_familia] = {'maes': [], 'filhos': []}
        
        dados_row['TIPO'] = '2. CAIXA MÃE'
        familias[ean_familia]['maes'].append(dados_row)
        
    # CASO 2: É UM FILHO
    if eh_filho:
        ean_familia = mapa_filho_para_mae[ean_atual] # O ID da família é o EAN da mãe dele
        
        if ean_familia not in familias:
            familias[ean_familia] = {'maes': [], 'filhos': []}
            
        dados_row['TIPO'] = '1. UNIDADE (FILHO)'
        familias[ean_familia]['filhos'].append(dados_row)

# ==============================================================================
# 4. GERAÇÃO DO RELATÓRIO
# ==============================================================================
lista_final = []
id_grupo_visual = 1

print(f"--> Encontrados {len(familias)} grupos familiares (completos ou incompletos).")

for ean_mae_key, dados in familias.items():
    
    lista_maes = dados['maes']
    lista_filhos = dados['filhos']
    
    # Define o status do grupo para ajudar na conferência
    status = ""
    if lista_maes and lista_filhos:
        status = "COMPLETO (Mãe e Filho presentes)"
    elif lista_maes and not lista_filhos:
        status = "INCOMPLETO (Só Mãe encontrada)"
    elif not lista_maes and lista_filhos:
        status = "INCOMPLETO (Só Filho encontrado)"
    
    # Adiciona Filhos primeiro (Unidades)
    for item in lista_filhos:
        item['ID_GRUPO'] = id_grupo_visual
        item['STATUS_FAMILIA'] = status
        item['EAN_MAE_REFERENCIA'] = ean_mae_key
        lista_final.append(item)
        
    # Adiciona Mães depois (Caixas)
    for item in lista_maes:
        item['ID_GRUPO'] = id_grupo_visual
        item['STATUS_FAMILIA'] = status
        item['EAN_MAE_REFERENCIA'] = ean_mae_key # É o próprio EAN dela
        lista_final.append(item)
    
    id_grupo_visual += 1

# ==============================================================================
# 5. SALVAR
# ==============================================================================
if not lista_final:
    print("\nNenhum produto relacionado foi encontrado no arquivo do banco.")
else:
    df_final = pd.DataFrame(lista_final)
    
    # Organiza colunas
    cols_order = ['ID_GRUPO', 'STATUS_FAMILIA', 'TIPO'] + COLUNAS_DESEJADAS + ['EAN_MAE_REFERENCIA']
    cols_order = [c for c in cols_order if c in df_final.columns]
    
    df_final = df_final[cols_order]
    
    print(f"\n[4/4] Gerando Excel: {NOME_RELATORIO}")
    try:
        df_final.to_excel(NOME_RELATORIO, index=False)
        print("✅ SUCESSO! Relatório gerado.")
        print("DICA: Verifique a coluna 'STATUS_FAMILIA' para ver quem está sozinho.")
    except Exception as e:
        print(f"Erro ao salvar: {e}")

input("\nPressione Enter para sair.")