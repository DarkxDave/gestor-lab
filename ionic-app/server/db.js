const mysql = require('mysql2/promise');
const dotenv = require('dotenv');
dotenv.config();

const pool = mysql.createPool({
  host: process.env.DB_HOST || 'localhost',
  user: process.env.DB_USER || 'root',
  password: process.env.DB_PASSWORD || '',
  database: process.env.DB_NAME || 'gestor_lab',
  port: process.env.DB_PORT ? parseInt(process.env.DB_PORT, 10) : 3306,
  waitForConnections: true,
  connectionLimit: 20,
  queueLimit: 0,
  enableKeepAlive: true,
  keepAliveInitialDelay: 0,
  connectTimeout: 10000,
  maxIdle: 10,
  idleTimeout: 60000
});

// Verificar conexión al iniciar
pool.getConnection()
  .then(conn => {
    console.log('✅ Conexión a MySQL establecida correctamente');
    console.log(`   Host: ${process.env.DB_HOST || 'localhost'}:${process.env.DB_PORT || 3306}`);
    console.log(`   Database: ${process.env.DB_NAME || 'gestor_lab'}`);
    conn.release();
  })
  .catch(err => {
    console.error('❌ Error al conectar a MySQL:', err.message);
    console.error('   Asegúrate de que:');
    console.error('   1. XAMPP/MySQL está corriendo (puerto 3306)');
    console.error('   2. La base de datos "gestor_lab" existe');
    console.error('   3. El archivo .env tiene la configuración correcta');
    console.error(`   Intentando conectar a: ${process.env.DB_HOST || 'localhost'}:${process.env.DB_PORT || 3306}`);
  });

module.exports = {
  pool,
  query: async (sql, params) => {
    try {
      const [rows] = await pool.execute(sql, params);
      return rows;
    } catch (error) {
      console.error('Error en consulta SQL:', error.message);
      throw error;
    }
  },
};
