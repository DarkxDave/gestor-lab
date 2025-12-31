const tpaModel = require('../models/tpaFormModel');

exports.save = async (req, res, next) => {
  try {
    const { sample_id } = req.body;
    if (!sample_id) {
      return res.status(400).json({ error: 'sample_id es requerido', message: 'sample_id es requerido' });
    }

    const payload = { ...req.body };
    Object.keys(req.body).forEach(k => {
      if (req.body[k] === 'on') payload[k] = true;
    });

    await tpaModel.save(sample_id, payload);
    const data = await tpaModel.getBySampleId(sample_id);
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
    
    // Timeout de 5 segundos para la consulta
    const timeoutPromise = new Promise((_, reject) => 
      setTimeout(() => reject(new Error('Timeout: la consulta tardó demasiado')), 5000)
    );
    
    const dataPromise = tpaModel.getBySampleId(sample_id);
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
