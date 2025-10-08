const express = require('express');
const router = express.Router();
const exportController = require('../controllers/exportController');

router.get('/excel', exportController.exportExcel);
router.get('/tpa-form', exportController.exportTPAForm);

module.exports = router;
