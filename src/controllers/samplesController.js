const samplesModel = require('../models/samples');

exports.list = async (req, res, next) => {
  try {
    const { q } = req.query;
    const samples = await samplesModel.list(q);
    res.render('samples', { title: 'Muestras', samples, q: q || '' });
  } catch (err) {
    next(err);
  }
};
