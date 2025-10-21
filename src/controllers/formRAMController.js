const ramModel = require('../models/formRAM');
const samples = require('../models/samples');

exports.renderForm = async (req, res) => {
  const sampleId = req.query.sample_id || '';
  res.render('formRAM', { title: 'Formulario RAM', data: null, message: null, sampleId });
};

exports.save = async (req, res, next) => {
  try {
    const { sample_id } = req.body;
    if (!sample_id) return res.status(400).render('formRAM', { title: 'Formulario RAM', data: null, message: 'sample_id es requerido', sampleId: '' });

    // Pasar todos los campos del body al modelo para persistencia
    await ramModel.save(sample_id, req.body);
    const data = await ramModel.getBySampleId(sample_id);
    res.render('formRAM', { title: 'Formulario RAM', data, message: 'Guardado correctamente', sampleId: sample_id });
  } catch (err) {
    next(err);
  }
};

exports.loadBySampleId = async (req, res, next) => {
  try {
    const { sample_id } = req.query;
    if (!sample_id) return res.render('formRAM', { title: 'Formulario RAM', data: null, message: 'Ingrese sample_id para cargar', sampleId: '' });
    const data = await ramModel.getBySampleId(sample_id);
    if (!data) return res.render('formRAM', { title: 'Formulario RAM', data: null, message: 'No encontrado', sampleId: sample_id });
    res.render('formRAM', { title: 'Formulario RAM', data, message: null, sampleId: sample_id });
  } catch (err) {
    next(err);
  }
};

// Guardar en todas las pestañas: esta versión mínima asegura presencia del sample en todas las tablas
exports.saveAll = async (req, res, next) => {
  try {
    const { sample_id } = req.body;
    if (!sample_id) return res.status(400).json({ ok: false, error: 'sample_id requerido' });
    await samples.ensureSample(sample_id);
    // Upserts mínimos en todas las tablas conocidas (solo aseguran fila)
    const upserts = [
      `INSERT INTO form_tpa_entries (sample_id, created_at, updated_at) VALUES (?, NOW(), NOW()) ON DUPLICATE KEY UPDATE updated_at=NOW()` ,
      `INSERT INTO form_ram_entries (sample_id, created_at, updated_at) VALUES (?, NOW(), NOW()) ON DUPLICATE KEY UPDATE updated_at=NOW()` ,
      `INSERT INTO form_rmyl_entries (sample_id, created_at, updated_at) VALUES (?, NOW(), NOW()) ON DUPLICATE KEY UPDATE updated_at=NOW()` ,
      `INSERT INTO form_ctcfe_entries (sample_id, created_at, updated_at) VALUES (?, NOW(), NOW()) ON DUPLICATE KEY UPDATE updated_at=NOW()` ,
      `INSERT INTO form_sal_entries (sample_id, created_at, updated_at) VALUES (?, NOW(), NOW()) ON DUPLICATE KEY UPDATE updated_at=NOW()` ,
      `INSERT INTO form_entero_entries (sample_id, created_at, updated_at) VALUES (?, NOW(), NOW()) ON DUPLICATE KEY UPDATE updated_at=NOW()` ,
      `INSERT INTO form_saureus_entries (sample_id, created_at, updated_at) VALUES (?, NOW(), NOW()) ON DUPLICATE KEY UPDATE updated_at=NOW()` ,
    ];
    const { query } = require('../db');
    for (const sql of upserts) {
      await query(sql, [sample_id]);
    }
    res.json({ ok: true });
  } catch (err) {
    next(err);
  }
};
