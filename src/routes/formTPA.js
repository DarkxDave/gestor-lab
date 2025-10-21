const express = require('express');
const router = express.Router();
const tpaController = require('../controllers/formTPAController');

router.get('/', async (req, res, next) => {
  const sample_id = req.query.sample_id || '';
  if (!sample_id) return tpaController.renderForm(req, res);
  req.query.sample_id = sample_id;
  return tpaController.loadBySampleId(req, res, next);
});

router.post('/save', (req, res, next) => tpaController.save(req, res, next));

module.exports = router;
