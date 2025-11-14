const { query } = require('../db');
const samples = require('./sampleModel');

exports.save = async (sample_id, { notes } = {}) => {
  await samples.ensureSample(sample_id);
  await query(
    `INSERT INTO form_ctcfe_entries (sample_id, notes, created_at, updated_at)
     VALUES (?, ?, NOW(), NOW())
     ON DUPLICATE KEY UPDATE notes=VALUES(notes), updated_at=NOW()` ,
    [sample_id, notes || null]
  );
};

exports.getBySampleId = async (sample_id) => {
  const rows = await query('SELECT * FROM form_ctcfe_entries WHERE sample_id=?', [sample_id]);
  return rows[0] || null;
};
