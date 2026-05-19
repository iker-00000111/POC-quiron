const express = require('express');
const cors = require('cors');
const sql = require('mssql');
const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());

// ==========================================
// 1. CONFIGURACIÓN DE AZURE SQL
// ==========================================
const dbConfig = {
    server: process.env.DB_SERVER,
    database: process.env.DB_NAME,
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    options: {
        encrypt: true, // Obligatorio para Azure SQL
        trustServerCertificate: false
    }
};

// ==========================================
// 2. MIDDLEWARE PARA VALIDAR EL TOKEN DE ENTRA ID
// ==========================================
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

    // Validar el token contra las claves públicas de Microsoft
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
        // Guardamos los datos del usuario autenticado en la petición por si hiciesen falta
        req.usuario = decoded; 
        next();
    });
};

// ==========================================
// 3. ENDPOINTS DE LA API
// ==========================================

// Endpoint protegido que heredará el flujo de la POC
app.get('/api/datos', validarTokenEntraID, async (req, res) => {
    try {
        // Abrir conexión con Azure SQL
        let pool = await sql.connect(dbConfig);
        
        // Realizar la consulta a la tabla que creamos en el portal
        let result = await pool.request().query('SELECT * FROM Clientes');
        
        // Responder al frontend con los datos obtenidos
        res.json(result.recordset);
    } catch (error) {
        console.error('Error al consultar Azure SQL:', error);
        res.status(500).json({ error: 'Error interno al consultar la base de datos.' });
    } finally {
        // Cerramos la conexión para no saturar el pool en la POC
        await sql.close();
    }
});

// Endpoint público de salud para comprobar que la API responde
app.get('/health', (req, res) => {
    res.json({ status: 'API levantada y corriendo perfectamente' });
});

// Levantar el servidor
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Servidor backend corriendo en http://localhost:${PORT}`);
});