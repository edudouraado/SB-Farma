const express = require('express');
const mysql = require('mysql');
const cors = require('cors'); 

const app = express();
const port = 3000;

// --- Configuração do MySQL (AJUSTE AS CREDENCIAIS) ---
const db = mysql.createConnection({
    host: 'localhost',      // Seu Host do MySQL (ou 127.0.0.1)
    user: 'root',           // Seu DB_USER
    password: '12345',      // Sua Senha
    database: 'inventario_db' // Nome do seu Banco de Dados
});

db.connect(err => {
    if (err) {
        console.error('Erro ao conectar ao banco de dados:', err);
        return;
    }
    console.log('Conectado ao MySQL com sucesso!');
});

// --- Middleware ---
app.use(cors()); // Permite que o Frontend (seu HTML) acesse a API
app.use(express.json());

// --- Endpoint de Busca: GET /api/produto/[codigo] ---
app.get('/api/produto/:codigo', (req, res) => {
    const codigoBarras = req.params.codigo;

    // A CONSULTA SQL ESSENCIAL:
    const sql = `
        SELECT nome_produto 
        FROM Tabela_Produtos 
        WHERE codigo_barras = ?
    `;

    db.query(sql, [codigoBarras], (err, result) => {
        if (err) {
            console.error('Erro na consulta SQL:', err);
            return res.status(500).send({ message: 'Erro interno do servidor' });
        }

        if (result.length > 0) {
            // Retorna o nome do produto encontrado
            res.json({ nome: result[0].nome_produto });
        } else {
            // Produto não encontrado
            res.status(404).send({ nome: 'Produto Não Encontrado' });
        }
    });
});


// Inicializa o servidor
app.listen(port, () => {
    console.log(`Servidor rodando em http://localhost:${port}`);
});