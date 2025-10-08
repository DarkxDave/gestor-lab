const express = require('express');
const router = express.Router();
const formAController = require('../controllers/formAController');

router.get('/', async (req, res, next) => {
  const sample_id = req.query.sample_id || '';
  if (!sample_id) return formAController.renderForm(req, res);
  req.query.sample_id = sample_id;
  return formAController.loadBySampleId(req, res, next);
});

router.post('/save', (req, res, next) => formAController.save(req, res, next));

module.exports = router;
