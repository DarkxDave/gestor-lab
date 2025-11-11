const ExcelJS = require('exceljs');
const tpaModel = require('../models/formTPA');
const ramModel = require('../models/formRAM');

function asCheck(v) {
  return (v === true || v === 1 || v === '1') ? '✓' : '';
}

function getTPASchema() {
  return [
    { header: 'sample_id', key: 'sample_id', width: 16, group: 'Identificación' },
    { header: 'Freezer 33-M', key: 'storage_freezer_33m', width: 14, group: 'Almacenamiento' },
    { header: 'Refrigerador 33-M', key: 'storage_refrigerador_33m', width: 16, group: 'Almacenamiento' },
    { header: 'Mesón siembra', key: 'storage_meson_siembra', width: 14, group: 'Almacenamiento' },
    { header: 'Gabinete Traspaso', key: 'storage_gabinete_traspaso', width: 16, group: 'Almacenamiento' },
    { header: 'Observaciones', key: 'observaciones', width: 30, group: 'Almacenamiento' },
  ];
}

async function buildTPASheet(wb) {
  const rows = await tpaModel.listAll();
  const ws = wb.addWorksheet('TPA');
  const schema = getTPASchema();
  ws.columns = schema.map(c => ({ key: c.key, width: c.width || 12 }));
  const totalCols = schema.length;
  ws.mergeCells(1, 1, 1, totalCols);
  const titleCell = ws.getCell(1, 1);
  titleCell.value = 'Formulario de Trazabilidad (TPA) - Resumen';
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  titleCell.font = { bold: true, size: 14 };
  // Group headers
  const groups = {}; const order = [];
  schema.forEach((c, i) => { const g = c.group || ''; if (!groups[g]) { groups[g] = { start: i+1, end: i+1 }; order.push(g);} else groups[g].end = i+1; });
  order.forEach(g => { const { start, end } = groups[g]; ws.mergeCells(2, start, 2, end); const cell = ws.getCell(2, start); cell.value = g; cell.alignment = { horizontal:'center' }; cell.font = { bold:true }; });
  ws.getRow(3).values = [, ...schema.map(c => c.header)]; ws.getRow(3).font = { bold: true };
  rows.forEach(r => {
    const ordered = schema.map(c => {
      const v = r[c.key];
      const keysBool = ['storage_freezer_33m','storage_refrigerador_33m','storage_meson_siembra','storage_gabinete_traspaso'];
      if (keysBool.includes(c.key)) return asCheck(v);
      return v ?? '';
    });
    ws.addRow(ordered);
  });
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 3 }];
}

async function buildRAMSheet(wb) {
  const rows = await ramModel.listAll();
  const ws = wb.addWorksheet('RAM');
  const schema = [
    { header: 'sample_id', key: 'sample_id', width: 16, group: 'Identificación' },
    // Fechas y análisis
    { header: 'Inicio Fecha', key: 'inicio_incubacion_fecha', width: 12, group: 'Fechas' },
    { header: 'Inicio Hora', key: 'inicio_incubacion_hora', width: 10, group: 'Fechas' },
    { header: 'Inicio Analista', key: 'inicio_incubacion_analista', width: 18, group: 'Fechas' },
    { header: 'Término Fecha', key: 'termino_analisis_fecha', width: 12, group: 'Fechas' },
    { header: 'Término Hora', key: 'termino_analisis_hora', width: 10, group: 'Fechas' },
    { header: 'Término Analista', key: 'termino_analisis_analista', width: 18, group: 'Fechas' },
    // Set de Análisis (cc2_*)
    { header: 'Pesado T°', key: 'cc2_pesado_temp', width: 12, group: 'Set de Análisis' },
    { header: 'Pesado UFC', key: 'cc2_pesado_ufc', width: 12, group: 'Set de Análisis' },
    { header: 'Siembra', key: 'cc2_siembra', width: 16, group: 'Set de Análisis' },
    { header: 'Hora inicio', key: 'cc2_hora_inicio', width: 12, group: 'Set de Análisis' },
    { header: 'Hora término', key: 'cc2_hora_termino', width: 12, group: 'Set de Análisis' },
    { header: 'T°', key: 'cc2_temp', width: 8, group: 'Set de Análisis' },
    { header: 'E.coli UFC', key: 'cc2_ecoli_ufc', width: 12, group: 'Set de Análisis' },
    { header: 'Blanco UFC', key: 'cc2_blanco_ufc', width: 12, group: 'Set de Análisis' },
    // Siembra
    { header: '<15 min', key: 'siembra_tiempo_ok', width: 10, group: 'Siembra' },
    { header: 'N° 10g/90ml', key: 'siembra_n_muestra_10g_90ml', width: 14, group: 'Siembra' },
    { header: 'N° 50g/450ml', key: 'siembra_n_muestra_50g_450ml', width: 16, group: 'Siembra' },
    // MIC
    { header: 'Desfavorable SI', key: 'mic_desfavorable_si', width: 14, group: 'MIC' },
    { header: 'Desfavorable NO', key: 'mic_desfavorable_no', width: 14, group: 'MIC' },
    { header: 'Tabla/Página', key: 'mic_tabla_pagina', width: 14, group: 'MIC' },
    { header: 'Límite', key: 'mic_limite', width: 12, group: 'MIC' },
    { header: 'Fecha Entrega', key: 'mic_fecha_entrega', width: 14, group: 'MIC' },
    { header: 'Hora Entrega', key: 'mic_hora_entrega', width: 12, group: 'MIC' },
    // Datos
    { header: 'Suspensión 1/den', key: 'datos_suspension_inicial_den', width: 14, group: 'Datos' },
    { header: 'Volumen [mL]', key: 'datos_volumen_petri_ml', width: 12, group: 'Datos' },
    // (Se removió la sección de Muestrario)
    // Notas
    { header: 'Notas', key: 'notes', width: 24, group: 'Notas' },
    { header: 'Observaciones', key: 'observaciones', width: 24, group: 'Notas' },
  ];
  ws.columns = schema.map(c => ({ key: c.key, width: c.width || 12 }));
  const totalCols = schema.length;
  ws.mergeCells(1, 1, 1, totalCols);
  const titleCell = ws.getCell(1, 1);
  titleCell.value = 'RAM - Resumen';
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  titleCell.font = { bold: true, size: 14 };
  // Group headers
  const groups = {}; const order = [];
  schema.forEach((c, i) => { const g = c.group || ''; if (!groups[g]) { groups[g] = { start: i+1, end: i+1 }; order.push(g);} else groups[g].end = i+1; });
  order.forEach(g => { const { start, end } = groups[g]; ws.mergeCells(2, start, 2, end); const cell = ws.getCell(2, start); cell.value = g; cell.alignment = { horizontal:'center' }; cell.font = { bold:true }; });
  ws.getRow(3).values = [, ...schema.map(c => c.header)]; ws.getRow(3).font = { bold: true };
  rows.forEach(r => {
    const ordered = schema.map(c => {
      const v = r[c.key];
      const boolish = new Set(['siembra_tiempo_ok','mic_desfavorable_si','mic_desfavorable_no']);
      if (boolish.has(c.key)) return asCheck(v);
      return v ?? '';
    });
    ws.addRow(ordered);
  });
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 3 }];
}

exports.exportAll = async (req, res, next) => {
  try {
    const sample_id = req.query.sample_id || null;
    const wb = new ExcelJS.Workbook();
    if (sample_id) {
      // Per-sample workbook: use the pretty TPA layout helper
      await require('./exportTPAController').addPrettySheetForSample(wb, sample_id);
      // RAM detailed sheet via controller helper
      const ramCtrl = require('./exportRAMController');
      await ramCtrl.addSheetForSample(wb, sample_id);
      await ramCtrl.addProvisorioSheetForSample(wb, sample_id);
    } else {
      // Summary workbook: all samples multi-sheet
      await buildTPASheet(wb);
      await buildRAMSheet(wb);
    }
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${sample_id ? 'muestra_'+sample_id : 'muestras_resumen'}.xlsx"`);
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    next(err);
  }
};
