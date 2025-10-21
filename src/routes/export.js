const express = require('express');
const router = express.Router();
const exportTPAController = require('../controllers/exportTPAController');
const exportRAMController = require('../controllers/exportRAMController');
const exportEnteroController = require('../controllers/exportEnteroController');
const exportCTCFEController = require('../controllers/exportCTCFEController');
const exportRMyLController = require('../controllers/exportRMyLController');
const exportSalController = require('../controllers/exportSalController');
const exportSaureusController = require('../controllers/exportSaureusController');
const exportMultiController = require('../controllers/exportMultiController');

router.get('/excel', exportTPAController.exportExcel);
router.get('/excel-all', exportMultiController.exportAll);
router.get('/tpa-form', exportTPAController.exportTPAForm);
router.get('/ram-form', exportRAMController.exportRAMForm);
router.get('/entero-form', exportEnteroController.exportEnteroForm);
router.get('/ctcfe-form', exportCTCFEController.exportCTCFEForm);
router.get('/rmyl-form', exportRMyLController.exportRMyLForm);
router.get('/sal-form', exportSalController.exportSalForm);
router.get('/saureus-form', exportSaureusController.exportSaureusForm);

module.exports = router;
