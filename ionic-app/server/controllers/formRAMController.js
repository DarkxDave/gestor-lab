const ramModel = require('../models/ramFormModel');

exports.save = async (req, res, next) => {
  try {
    const { sample_id } = req.body;
    if (!sample_id) {
      return res.status(400).json({ error: 'sample_id es requerido', message: 'sample_id es requerido' });
    }

    await ramModel.save(sample_id, req.body);
    const data = await ramModel.getBySampleId(sample_id);
    res.json({ data, message: 'Guardado correctamente' });
  } catch (err) {
    next(err);
  }
};

exports.loadBySampleId = async (req, res, next) => {
  try {
    const { sample_id } = req.query;
    if (!sample_id) {
      return res.status(400).json({ error: 'Ingrese sample_id para cargar', message: 'Ingrese sample_id para cargar' });
    }
    
    const timeoutPromise = new Promise((_, reject) => 
      setTimeout(() => reject(new Error('Timeout: la consulta tardó demasiado')), 5000)
    );
    
    const dataPromise = ramModel.getBySampleId(sample_id);
    const data = await Promise.race([dataPromise, timeoutPromise]);
    
    if (!data) {
      return res.status(404).json({ error: 'No encontrado', message: 'No encontrado', data: null });
    }
    res.json({ data, message: null });
  } catch (err) {
    console.error('Error en loadBySampleId:', err.message);
    if (err.message.includes('Timeout')) {
      return res.status(504).json({ error: 'Timeout', message: 'La consulta tardó demasiado tiempo' });
    }
    next(err);
  }
};

// Guardar en todas las pestañas
exports.saveAll = async (req, res, next) => {
  try {
    const { sample_id } = req.body;
    if (!sample_id) return res.status(400).json({ ok: false, error: 'sample_id requerido' });
    await samples.ensureSample(sample_id);
    // Asegurar filas mínimas vía modelos correspondientes
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
