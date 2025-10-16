import pandas as pd
import numpy as np
import calendar
import os

# --- CONFIGURAÇÃO ---
file_name = "Estudo de Perdas dos Vencidos e Avarias.xlsx" 
output_file = "Perdas_Diarias_Rateadas_SOLUCAO_FINAL.csv"
DELIMITADOR_SAIDA = ';'
ANO_REFERENCIA = 2025 
NOME_DA_ABA = 'Planilha1' # <--- CONFIRME SE ESTE É O NOME CORRETO DA ABA NO SEU EXCEL!

# Dicionário para mapear nomes de meses para números de mês
MESES = {
    'Janeiro': 1, 'Fevereiro': 2, 'Março': 3, 'Abril': 4, 'Maio': 5,
    'Junho': 6, 'Julho': 7, 'Agosto': 8, 'Setembro': 9, 'Outubro': 10,
    'Novembro': 11, 'Dezembro': 12
}

def processar_rateio_diario():
    try:
        print(f"Lendo dados do arquivo XLSX: {file_name}")
        
        # 1. Carregar os dados (USANDO EXCEL e forçando a vírgula como decimal)
        # O parâmetro thousands='.' e decimal=',' é a tentativa final de convencer o Pandas
        # a ler o formato brasileiro corretamente.
        df = pd.read_excel(
            file_name, 
            sheet_name=NOME_DA_ABA, 
            engine='openpyxl',
            thousands='.',  # Diz que o ponto é separador de milhar
            decimal=','     # Diz que a vírgula é o separador decimal
        )
        
        # 2. Saneamento e Limpeza de Cabeçalhos
        df.columns = df.columns.astype(str).str.strip().str.replace('‡', 'ç', regex=False)
        df.columns = df.columns.str.replace('  - ', ' - ', regex=False)
        
        perda_cols = [col for col in df.columns if col not in ['Lojas'] and 'Vd Lqd' not in col]
        
        # 3. Remover o Pivot (Melt) - Focando apenas nas perdas
        df_perdas = df.melt(id_vars='Lojas', value_vars=perda_cols, var_name='Mês_Nome', value_name='Perda_Mensal')

        # 4. Limpeza (Apenas removendo linhas vazias, a conversão deve ter sido automática)
        df_perdas = df_perdas.dropna(subset=['Perda_Mensal']).copy()

        # 5. Cálculo do Rateio Diário
        df_perdas['Mês_Nome'] = df_perdas['Mês_Nome'].str.strip()
        df_perdas['Mês'] = df_perdas['Mês_Nome'].map(MESES)
        df_perdas['Ano'] = ANO_REFERENCIA
        
        df_perdas = df_perdas.dropna(subset=['Mês']).copy()
        
        df_perdas['Mês'] = df_perdas['Mês'].astype(int)
        df_perdas['Ano'] = df_perdas['Ano'].astype(int)

        df_perdas['Dias_No_Mês'] = df_perdas.apply(
            lambda row: calendar.monthrange(row['Ano'], row['Mês'])[1], axis=1
        )
        
        df_perdas['Perda_Diaria_Rateada'] = df_perdas['Perda_Mensal'] / df_perdas['Dias_No_Mês']

        # 6. EXPANDIR PARA LINHAS DIÁRIAS (Criar uma linha para cada dia)
        lista_final = []
        
        for index, row in df_perdas.iterrows():
            loja = row['Lojas']
            ano = row['Ano']
            mes = row['Mês']
            dias = row['Dias_No_Mês']
            perda_diaria = row['Perda_Diaria_Rateada']
            
            for dia in range(1, dias + 1):
                data = pd.to_datetime(f"{ano}-{mes}-{dia}")
                
                lista_final.append({
                    'Data': data.strftime('%d/%m/%Y'),
                    'Loja': loja,
                    'Perda_Diaria_Rateada': perda_diaria
                })
        
        df_final = pd.DataFrame(lista_final)

        # 7. Salvar o Resultado
        df_final.to_csv(output_file, index=False, sep=DELIMITADOR_SAIDA, encoding='utf-8')

        print("\n--- SUCESSO NA TRANSFORMAÇÃO E RATEIO ---")
        print(f"O novo arquivo '{output_file}' foi gerado e está pronto para o Power BI.")
        
        # CONFERÊNCIA FINAL
        loja_01_jan = df_perdas[(df_perdas['Lojas'] == 'Loja 01') & (df_perdas['Mês_Nome'] == 'Janeiro')]
        perda_total_jan = loja_01_jan['Perda_Mensal'].iloc[0]
        perda_diaria_esperada = loja_01_jan['Perda_Diaria_Rateada'].iloc[0]
        
        print(f"\nCONFERÊNCIA (Loja 01, Janeiro):")
        print(f"Perda Total Lida: R$ {perda_total_jan:.2f}")
        print(f"Valor Diário Rateado (DEVE SER 123.07): R$ {perda_diaria_esperada:.4f}")
        
    except FileNotFoundError:
        print(f"\n--- ERRO: Arquivo Não Encontrado ---")
        print(f"Certifique-se de que o arquivo '{file_name}' está no mesmo diretório do script.")
    except Exception as e:
        print(f"Ocorreu um erro durante o processamento: {e}")

# Executa a função
processar_rateio_diario()