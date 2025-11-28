const ramModel = require('../models/ramFormModel');
const samples = require('../models/sampleModel');

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
    // Asegurar filas mínimas vía modelos (sin SQL directo en controlador)
    const tpa = require('../models/tpaFormModel');
    const ram = require('../models/ramFormModel');
    const rmyl = require('../models/rmylFormModel');
    const ctcfe = require('../models/ctcfeFormModel');
    const sal = require('../models/salFormModel');
    const entero = require('../models/enteroFormModel');
    const saureus = require('../models/saureusFormModel');
    await Promise.all([
      tpa.ensureRow(sample_id),
      ram.ensureRow(sample_id),
      rmyl.ensureRow(sample_id),
      ctcfe.ensureRow(sample_id),
      sal.ensureRow(sample_id),
      entero.ensureRow(sample_id),
      saureus.ensureRow(sample_id),
    ]);
    res.json({ ok: true });
  } catch (err) {
    next(err);
  }
};
