const express = require('express');
const router = express.Router();

function render(view){
  return (req, res) => {
    const sampleId = req.query.sample_id || '';
    res.render(view, { title: view, sampleId });
  };
}

router.get('/form-ram', render('formRAM'));
router.get('/form-rmyl', render('formRMyL'));
router.get('/form-ctcfe', render('formCTCFEcoli'));
router.get('/form-sal', render('formSal'));
router.get('/form-entero', render('formEntero'));
router.get('/form-saureus', render('formSaureus'));

module.exports = router;
