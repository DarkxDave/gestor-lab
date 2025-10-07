const express = require('express');
const router = express.Router();
const exportController = require('../controllers/exportController');

router.get('/excel', exportController.exportExcel);

module.exports = router;
