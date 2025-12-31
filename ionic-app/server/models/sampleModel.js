const { query } = require('../db');

/**
 * Asegura que exista una muestra en la base de datos.
 * Si el sample_id ya existe, no hace nada (INSERT IGNORE).
 */
exports.ensureSample = async (sample_id) => {
  await query('INSERT IGNORE INTO samples (sample_id) VALUES (?)', [sample_id]);
};

/**
 * Lista todas las muestras, opcionalmente filtrando por término de búsqueda.
 */
exports.list = async (searchTerm) => {
  if (searchTerm) {
    return await query(
      'SELECT id, sample_id FROM samples WHERE sample_id LIKE ? ORDER BY id ASC LIMIT 200',
      [`%${searchTerm}%`]
    );
  }
  return await query(
    'SELECT id, sample_id FROM samples ORDER BY id ASC LIMIT 200'
  );
};
