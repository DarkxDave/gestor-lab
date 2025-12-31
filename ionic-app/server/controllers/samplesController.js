const samplesModel = require('../models/sampleModel');

/**
 * Lista todas las muestras con búsqueda opcional.
 * @param {Object} req - Request con query param searchTerm opcional
 * @param {Object} res - Response con JSON array de muestras
 * @param {Function} next - Middleware de manejo de errores
 */
exports.list = async (req, res, next) => {
  try {
    const { searchTerm } = req.query;
    const samples = await samplesModel.list(searchTerm);
    res.json(samples);
  } catch (err) {
    console.error('Error fetching samples:', err);
    res.status(500).json({ error: 'Error fetching samples', message: err.message });
  }
};
