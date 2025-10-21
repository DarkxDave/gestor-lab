const express = require('express');
const router = express.Router();
const ctrl = require('../controllers/formEnteroController');

router.get('/', (req, res, next) => {
  const sample_id = req.query.sample_id || '';
  if (!sample_id) return ctrl.renderForm(req, res);
  req.query.sample_id = sample_id;
  return ctrl.loadBySampleId(req, res, next);
});

router.post('/save', ctrl.save);

module.exports = router;
