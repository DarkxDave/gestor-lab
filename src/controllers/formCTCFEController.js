const model = require('../models/formCTCFE');

exports.renderForm = async (req, res) => {
  const sampleId = req.query.sample_id || '';
  res.render('formCTCFEcoli', { title: 'Formulario CT, CF y E.coli', data: null, message: null, sampleId });
};

exports.save = async (req, res, next) => {
  try {
    const { sample_id, notes } = req.body;
    if (!sample_id) return res.status(400).render('formCTCFEcoli', { title: 'Formulario CT, CF y E.coli', data: null, message: 'sample_id es requerido', sampleId: '' });
    await model.save(sample_id, { notes });
    const data = await model.getBySampleId(sample_id);
    res.render('formCTCFEcoli', { title: 'Formulario CT, CF y E.coli', data, message: 'Guardado correctamente', sampleId: sample_id });
  } catch (err) { next(err); }
};

exports.loadBySampleId = async (req, res, next) => {
  try {
    const { sample_id } = req.query;
    if (!sample_id) return res.render('formCTCFEcoli', { title: 'Formulario CT, CF y E.coli', data: null, message: 'Ingrese sample_id para cargar', sampleId: '' });
    const data = await model.getBySampleId(sample_id);
    if (!data) return res.render('formCTCFEcoli', { title: 'Formulario CT, CF y E.coli', data: null, message: 'No encontrado', sampleId });
    res.render('formCTCFEcoli', { title: 'Formulario CT, CF y E.coli', data, message: null, sampleId });
  } catch (err) { next(err); }
};
