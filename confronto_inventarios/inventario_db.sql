USE inventario_db
SELECT * FROM inventario_mudancas;

SELECT
    descrição,
    Dif_Antes,
    Dif_Depois,
    REPLACE(SUBSTRING_INDEX(Loja, ' - ', 1), 'Loja ', '') AS N_Loja 
FROM
    inventario_mudancas
WHERE
    tipo_de_mudanca = 'MUDANÇA_DE_STATUS'

    AND Status_Antes <> 'Estoque_Ok'
    AND Status_Depois <> 'Estoque_Ok'

    AND REPLACE(SUBSTRING_INDEX(Loja, ' - ', 1), 'Loja ', '') = '10';