const fs = require('fs');
const path = require('path');
const mysql = require('mysql2/promise');
const dotenv = require('dotenv');
dotenv.config();

(async () => {
  const sqlPath = path.join(__dirname, 'init_db.sql');
  const sql = fs.readFileSync(sqlPath, 'utf8');
  const connection = await mysql.createConnection({
    host: process.env.DB_HOST || 'localhost',
    user: process.env.DB_USER || 'root',
    password: process.env.DB_PASSWORD || '',
    port: process.env.DB_PORT ? parseInt(process.env.DB_PORT, 10) : 3306,
    multipleStatements: true,
  });
  try {
    await connection.query(sql);
    console.log('Base de datos inicializada.');
  } finally {
    await connection.end();
  }
})().catch(err => {
  console.error('Error al inicializar la base de datos:', err.message);
  process.exit(1);
});
