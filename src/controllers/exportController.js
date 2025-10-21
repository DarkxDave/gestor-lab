// DEPRECATED: TPA export logic moved to src/controllers/exportTPAController.js
// This file intentionally left as a stub to avoid breaking imports.

exports.exportExcel = (req, res) => {
  res.status(410).send('exportController.exportExcel está obsoleto. Usa exportTPAController.exportExcel');
};

exports.exportTPAForm = (req, res) => {
  res.status(410).send('exportController.exportTPAForm está obsoleto. Usa exportTPAController.exportTPAForm');
};
