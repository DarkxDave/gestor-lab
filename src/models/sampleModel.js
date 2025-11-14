const { query } = require('../db');

exports.ensureSample = async (sample_id) => {
  await query('INSERT IGNORE INTO samples (sample_id) VALUES (?)', [sample_id]);
};

exports.list = async (q) => {
  if (q) {
    return await query('SELECT id, sample_id FROM samples WHERE sample_id LIKE ? ORDER BY id DESC LIMIT 200', [`%${q}%`]);
  }
  return await query('SELECT id, sample_id FROM samples ORDER BY id DESC LIMIT 200');
};
