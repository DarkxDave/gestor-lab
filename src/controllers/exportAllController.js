const ExcelJS = require('exceljs');
// NOTA: Este controlador actúa como el "hub" de exportación multi‑pestaña por muestra.
// Aquí se orquestan las hojas TPA, RAM y las pestañas simples (CTCFE, Entero, RM y L, Sal, Saureus).
// Para agregar una nueva pestaña por muestra, impórtela y súmela en la secuencia de abajo
// (idealmente exponiendo un método addSheetForSample en su controlador específico).
const exportTPA = require('./exportTPAController');
const exportRAM = require('./exportRAMController');
const modelCTCFE = require('../models/ctcfeFormModel');
const modelEntero = require('../models/enteroFormModel');
const modelRMyL = require('../models/rmylFormModel');
const modelSal = require('../models/salFormModel');
const modelSaureus = require('../models/saureusFormModel');

function addSimpleNotesSheet(wb, sheetName, title, sample_id, notes) {
  const ws = wb.addWorksheet(sheetName);
  ws.mergeCells('B1:T1'); ws.getCell('B1').value = title; ws.getCell('B1').alignment = { horizontal:'center' }; ws.getCell('B1').font = { bold:true, size:16 };
  ws.mergeCells('I3:K3'); const idCell = ws.getCell('I3'); idCell.value = `${sample_id}`; idCell.alignment = { horizontal:'center', vertical:'middle' }; idCell.font = { bold:true, size:14 }; idCell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFF2F2F2'} };
  ws.getCell('H3').value = 'ALI'; ws.getCell('H3').alignment = { horizontal:'center', vertical:'middle' };
  ws.getColumn(2).width = 28; ws.getColumn(3).width = 90;
  ws.getCell('B5').value = 'Notas'; ws.getCell('B5').font = { bold:true };
  ws.mergeCells('C5:T9'); ws.getCell('C5').value = notes || ''; ws.getCell('C5').alignment = { vertical:'top', wrapText:true };
}

exports.exportAllForms = async (req, res, next) => {
  try {
    const sample_id = req.query.sample_id;
    if (!sample_id) return res.status(400).send('Parámetro sample_id requerido. Ej: /export/all-forms?sample_id=XYZ');

    const wb = new ExcelJS.Workbook();

    // 1) TPA (unificado al mismo helper que RAM-style)
    await exportTPA.addSheetForSample(wb, sample_id);
    // 2) RAM detallado con fórmulas
    await exportRAM.addSheetForSample(wb, sample_id);

    // 3) Hojas simples (mientras no exista "addSheetForSample" propio)
    const ctcfe = await modelCTCFE.getBySampleId(sample_id);
    addSimpleNotesSheet(wb, 'CTCFE', 'TRAZABILIDAD Y ANÁLISIS - CT, CF y E.coli', sample_id, ctcfe && ctcfe.notes);

    const entero = await modelEntero.getBySampleId(sample_id);
    addSimpleNotesSheet(wb, 'Entero', 'TRAZABILIDAD Y ANÁLISIS - ENTERO', sample_id, entero && entero.notes);

    const rmyl = await modelRMyL.getBySampleId(sample_id);
    addSimpleNotesSheet(wb, 'RM y L', 'TRAZABILIDAD Y ANÁLISIS - RM y L', sample_id, rmyl && rmyl.notes);

    const sal = await modelSal.getBySampleId(sample_id);
    addSimpleNotesSheet(wb, 'Sal', 'TRAZABILIDAD Y ANÁLISIS - Sal', sample_id, sal && sal.notes);

    const saureus = await modelSaureus.getBySampleId(sample_id);
    addSimpleNotesSheet(wb, 'Saureus', 'TRAZABILIDAD Y ANÁLISIS - Saureus', sample_id, saureus && saureus.notes);

    // Entrega del archivo resultante
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="muestra_${sample_id}.xlsx"`);
    await wb.xlsx.write(res); res.end();
  } catch (err) { next(err); }
};

/**
 * Cómo agregar una nueva hoja (p.ej. "HongoTorula"):
 * 1. Cree o exponga una función addSheetForSample(wb, sample_id, data) en su controlador específico,
 *    que añada la hoja al workbook recibido (sin escribir la respuesta).
 * 2. Impórtela aquí y añádala a la secuencia (como TPA/RAM), o
 * 3. Temporalmente, use addSimpleNotesSheet para mostrar un encabezado y "Notas" desde su modelo.
 *
 * Ejemplo (simple):
 *   const modelHongo = require('../models/formHongoTorula');
 *   const hongo = await modelHongo.getBySampleId(sample_id);
 *   addSimpleNotesSheet(wb, 'HongoTorula', 'TRAZABILIDAD Y ANÁLISIS - HongoTorula', sample_id, hongo && hongo.notes);
 *
 * Ejemplo (avanzado):
 *   const exportHongo = require('./exportHongoTorulaController');
 *   await exportHongo.addSheetForSample(wb, sample_id);
 */
