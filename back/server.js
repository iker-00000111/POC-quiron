const express = require('express');
const cors = require('cors');
const sql = require('mssql');
const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');
require('dotenv').config();

const app = express();
app.use(cors({
  origin: process.env.CORS_ORIGIN || "https://blue-pond-0fec4b703.7.azurestaticapps.net"
}));
app.use(express.json());
app.use((req, res, next) => {
  res.setHeader('Cross-Origin-Opener-Policy', 'same-origin-allow-popups');
  next();
});

const dbConfig = {
    server: process.env.DB_SERVER,
    database: process.env.DB_NAME,
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    options: {
        encrypt: true,
        trustServerCertificate: false
    }
};

const client = jwksClient({
    jwksUri: `https://login.microsoftonline.com/${process.env.TENANT_ID}/discovery/v2.0/keys`
});

function getKey(header, callback) {
    client.getSigningKey(header.kid, function (err, key) {
        if (err) return callback(err);
        const signingKey = key.publicKey || key.rsaPublicKey;
        callback(null, signingKey);
    });
}

const validarTokenEntraID = (req, res, next) => {
    const authHeader = req.headers.authorization;
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({ error: 'Acceso denegado. Token no proporcionado.' });
    }
    const token = authHeader.split(' ')[1];
    jwt.verify(token, getKey, {
        audience: [
            process.env.BACKEND_CLIENT_ID, 
            `api://${process.env.BACKEND_CLIENT_ID}`
        ],
    }, (err, decoded) => {
        if (err) {
            console.error("❌ Error de validación del Token:", err.message);
            return res.status(403).json({ error: 'Token inválido o expirado.' });
        }
        req.usuario = decoded; 
        next();
    });
};

app.get('/api/datos', validarTokenEntraID, async (req, res) => {
    try {
        let pool = await sql.connect(dbConfig);
        let result = await pool.request().query('SELECT * FROM Clientes');
        res.json(result.recordset);
    } catch (error) {
        console.error('Error al consultar Azure SQL:', error);
        res.status(500).json({ error: 'Error interno al consultar la base de datos.' });
    } finally {
        await sql.close();
    }
});

app.get('/health', (req, res) => {
    res.json({ status: 'API levantada y corriendo perfectamente' });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Servidor backend corriendo en http://localhost:${PORT}`);
});