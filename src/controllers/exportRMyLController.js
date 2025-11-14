const ExcelJS = require('exceljs');
const model = require('../models/rmylFormModel');

exports.exportRMyLForm = async (req, res, next) => {
  try {
    const sample_id = req.query.sample_id;
    if (!sample_id) return res.status(400).send('Parámetro sample_id requerido. Ej: /export/rmyl-form?sample_id=XYZ');
    const data = await model.getBySampleId(sample_id);
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('RM y L');
    ws.mergeCells('B1:T1'); ws.getCell('B1').value = 'TRAZABILIDAD Y ANÁLISIS - RM y L'; ws.getCell('B1').alignment = { horizontal:'center' }; ws.getCell('B1').font = { bold:true, size:16 };
    ws.mergeCells('I3:K3'); const idCell = ws.getCell('I3'); idCell.value = `${sample_id}`; idCell.alignment = { horizontal:'center', vertical:'middle' }; idCell.font = { bold:true, size:14 }; idCell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFF2F2F2'} };
    ws.getCell('H3').value = 'ALI'; ws.getCell('H3').alignment = { horizontal:'center', vertical:'middle' };
    ws.getColumn(2).width = 28; ws.getColumn(3).width = 90; ws.getCell('B5').value = 'Notas'; ws.getCell('B5').font = { bold:true }; ws.mergeCells('C5:T9'); ws.getCell('C5').value = (data && data.notes) || ''; ws.getCell('C5').alignment = { vertical:'top', wrapText:true };
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); res.setHeader('Content-Disposition', `attachment; filename="RMyL_${sample_id}.xlsx"`); await wb.xlsx.write(res); res.end();
  } catch (err) { next(err); }
};
