const formBModel = require('../models/formB');

exports.renderForm = async (req, res) => {
  res.render('formB', { title: 'Formulario B', data: null, message: null });
};

exports.save = async (req, res, next) => {
  try {
    const { sample_id, notes } = req.body;
    const approved = req.body.approved === 'on';
    const qc_pass = req.body.qc_pass === 'on';

    if (!sample_id) return res.status(400).render('formB', { title: 'Formulario B', data: null, message: 'sample_id es requerido' });

    await formBModel.save(sample_id, { notes, approved, qc_pass });
    const data = await formBModel.getBySampleId(sample_id);
    res.render('formB', { title: 'Formulario B', data, message: 'Guardado correctamente' });
  } catch (err) {
    next(err);
  }
};

exports.loadBySampleId = async (req, res, next) => {
  try {
    const { sample_id } = req.query;
    if (!sample_id) return res.render('formB', { title: 'Formulario B', data: null, message: 'Ingrese sample_id para cargar' });
    const data = await formBModel.getBySampleId(sample_id);
    if (!data) return res.render('formB', { title: 'Formulario B', data: null, message: 'No encontrado' });
    res.render('formB', { title: 'Formulario B', data, message: null });
  } catch (err) {
    next(err);
  }
};
