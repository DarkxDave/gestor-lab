const fs = require('fs');
const path = require('path');
const mysql = require('mysql2/promise');
const dotenv = require('dotenv');
dotenv.config();

(async () => {
  const migrationsDir = path.join(__dirname, 'migrations');
  const files = fs
    .readdirSync(migrationsDir)
    .filter(f => f.endsWith('.sql'))
    .sort();

  if (files.length === 0) {
    console.log('No hay migraciones para aplicar.');
    return;
  }

  const connection = await mysql.createConnection({
    host: process.env.DB_HOST || 'localhost',
    user: process.env.DB_USER || 'root',
    password: process.env.DB_PASSWORD || '',
    database: process.env.DB_NAME || 'gestor_lab',
    port: process.env.DB_PORT ? parseInt(process.env.DB_PORT, 10) : 3306,
    multipleStatements: true,
  });

  try {
    for (const file of files) {
      const full = path.join(migrationsDir, file);
      const sql = fs.readFileSync(full, 'utf8');
      if (!sql.trim()) continue;
      console.log(`Aplicando migración: ${file}`);
      try {
        await connection.query(sql);
        console.log(`OK: ${file}`);
      } catch (err) {
        const msg = String(err && err.message || '');
        // Ignora errores típicos si ya se aplicó: columna duplicada o inexistente
        if (/Duplicate column name/i.test(msg) || /can\'t drop column/i.test(msg) || /check that column\/key exists/i.test(msg)) {
          console.warn(`Aviso (ignorado): ${msg}`);
          continue;
        }
        throw err;
      }
    }
  } finally {
    await connection.end();
  }
})().catch(err => {
  console.error('Error al aplicar migraciones:', err.message);
  process.exit(1);
});
