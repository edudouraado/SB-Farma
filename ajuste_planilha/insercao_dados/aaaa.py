import mysql.connector
import pandas as pd

# --- CONFIGURAÇÕES ---
# Preencha com os dados do seu Aiven (pegue no MySQL Workbench ou no site da Aiven)
DB_CONFIG = {
    'host': 'prevencidos-sbfarma-edudouraado-pjc.j.aivencloud.com', # Sem a porta e sem https://
    'port': 14948,
    'user': 'avnadmin',
    'password': 'AVNS_P6JdPC4OZdgBwOZ1YE9', # Cole a senha do Aiven aqui
    'database': 'defaultdb',
    'ssl_ca': '', # Pode deixar vazio ou remover se der erro, mas no Aiven geralmente precisa de 'ssl_disabled' False
    'ssl_disabled': False 
}

ARQUIVO_CSV = 'produtos_completo_mysql.csv'

print("Lendo CSV...")
df = pd.read_csv(ARQUIVO_CSV)

# Substituir NaN (vazios) por None para o SQL não reclamar
df = df.where(pd.notnull(df), None)

print("Conectando ao banco...")
try:
    conn = mysql.connector.connect(**DB_CONFIG)
    cursor = conn.cursor()

    # Query de Inserção Otimizada
    sql = """
    INSERT INTO produtos (codigo_int, codigo_barras, nome, classificacao_geral)
    VALUES (%s, %s, %s, %s)
    ON DUPLICATE KEY UPDATE 
        nome = VALUES(nome), 
        classificacao_geral = VALUES(classificacao_geral),
        codigo_int = VALUES(codigo_int)
    """

    # Converte os dados para lista de tuplas
    val = [tuple(x) for x in df.values]

    print(f"Inserindo {len(val)} produtos... (Isso é rápido!)")
    
    # O executemany manda em pacotes grandes
    cursor.executemany(sql, val)
    conn.commit()

    print(f"SUCESSO! {cursor.rowcount} registros inseridos/atualizados.")

except Exception as e:
    print(f"Erro: {e}")

finally:
    if 'conn' in locals() and conn.is_connected():
        cursor.close()
        conn.close()
        print("Conexão fechada.")