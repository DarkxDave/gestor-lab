const { query } = require('../db');
const samples = require('./samples');

exports.save = async (sample_id, { notes, approved, qc_pass }) => {
  await samples.ensureSample(sample_id);
  const rows = await query('SELECT id FROM form_b_entries WHERE sample_id = ?', [sample_id]);
  if (rows.length) {
    await query('UPDATE form_b_entries SET notes=?, approved=?, qc_pass=?, updated_at=NOW() WHERE sample_id=?', [notes || null, !!approved, !!qc_pass, sample_id]);
  } else {
    await query('INSERT INTO form_b_entries (sample_id, notes, approved, qc_pass, created_at, updated_at) VALUES (?, ?, ?, ?, NOW(), NOW())', [sample_id, notes || null, !!approved, !!qc_pass]);
  }
};

exports.getBySampleId = async (sample_id) => {
  const rows = await query('SELECT * FROM form_b_entries WHERE sample_id = ?', [sample_id]);
  return rows[0] || null;
};
