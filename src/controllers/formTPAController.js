const tpaModel = require('../models/formTPA');

exports.renderForm = async (req, res) => {
  const sampleId = req.query.sample_id || '';
  res.render('formTPA', { title: 'Formulario de Trazabilidad (TPA)', data: null, message: null, sampleId });
};

exports.save = async (req, res, next) => {
  try {
    const { sample_id } = req.body;
    if (!sample_id) return res.status(400).render('formTPA', { title: 'Formulario de Trazabilidad (TPA)', data: null, message: 'sample_id es requerido', sampleId: '' });

    const payload = { ...req.body };
    Object.keys(req.body).forEach(k => {
      if (req.body[k] === 'on') payload[k] = true;
    });

    await tpaModel.save(sample_id, payload);
    const data = await tpaModel.getBySampleId(sample_id);
    res.render('formTPA', { title: 'Formulario de Trazabilidad (TPA)', data, message: 'Guardado correctamente', sampleId: sample_id });
  } catch (err) {
    next(err);
  }
};

exports.loadBySampleId = async (req, res, next) => {
  try {
    const { sample_id } = req.query;
    if (!sample_id) return res.render('formTPA', { title: 'Formulario de Trazabilidad (TPA)', data: null, message: 'Ingrese sample_id para cargar', sampleId: '' });
    const data = await tpaModel.getBySampleId(sample_id);
    if (!data) return res.render('formTPA', { title: 'Formulario de Trazabilidad (TPA)', data: null, message: 'No encontrado', sampleId: sample_id });
    res.render('formTPA', { title: 'Formulario de Trazabilidad (TPA)', data, message: null, sampleId: sample_id });
  } catch (err) {
    next(err);
  }
};
