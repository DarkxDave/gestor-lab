const express = require('express');
const router = express.Router();

// Backward compatibility: redirect legacy routes to new /form-tpa
router.get('/', (req, res) => res.redirect(302, '/form-tpa'));
router.post('/save', (req, res) => res.redirect(302, '/form-tpa'));
router.get('/load', (req, res) => {
	const q = req.query.sample_id ? `?sample_id=${encodeURIComponent(req.query.sample_id)}` : '';
	res.redirect(302, '/form-tpa' + q);
});

module.exports = router;
