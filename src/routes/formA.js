const express = require('express');
const router = express.Router();
const formAController = require('../controllers/formAController');

router.get('/', formAController.renderForm);
router.post('/save', formAController.save);
router.get('/load', formAController.loadBySampleId);

module.exports = router;
