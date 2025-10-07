const express = require('express');
const router = express.Router();
const formBController = require('../controllers/formBController');

router.get('/', formBController.renderForm);
router.post('/save', formBController.save);
router.get('/load', formBController.loadBySampleId);

module.exports = router;
