const express = require('express');
const router = express.Router();
const samplesController = require('../controllers/samplesController');

router.get('/', samplesController.list);

module.exports = router;
