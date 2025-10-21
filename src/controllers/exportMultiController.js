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
    // Control Ambiental
    { header: 'CA Pesado T°', key: 'ca_pesado_temp', width: 12, group: 'Control Ambiental' },
    { header: 'CA Pesado UFC', key: 'ca_pesado_ufc', width: 12, group: 'Control Ambiental' },
    { header: 'CA Siembra', key: 'ca_siembra', width: 18, group: 'Control Ambiental' },
    { header: 'CA E.coli UFC', key: 'ca_ecoli_ufc', width: 12, group: 'Control Ambiental' },
    { header: 'CA Blanco UFC', key: 'ca_blanco_ufc', width: 12, group: 'Control Ambiental' },
    // Siembra
    { header: '<15 min', key: 'siembra_tiempo_ok', width: 10, group: 'Siembra' },
    { header: 'Minutos', key: 'siembra_tiempo_minutos', width: 10, group: 'Siembra' },
    { header: 'N° 10g/90ml', key: 'siembra_n_muestra_10g_90ml', width: 14, group: 'Siembra' },
    { header: 'N° 50g/450ml', key: 'siembra_n_muestra_50g_450ml', width: 16, group: 'Siembra' },
    // CC
    { header: 'Dup ALI Detalle', key: 'cc_duplicado_ali_detalle', width: 18, group: 'CC' },
    { header: 'Dup ALI Análisis', key: 'cc_duplicado_ali_analisis', width: 18, group: 'CC' },
    { header: 'Dup ALI Cumple', key: 'cc_duplicado_ali_cumple', width: 12, group: 'CC' },
    { header: '(+) y Blanco Det.', key: 'cc_control_pos_blanco_ali_detalle', width: 18, group: 'CC' },
    { header: '(+) y Blanco Anál.', key: 'cc_control_pos_blanco_ali_analisis', width: 18, group: 'CC' },
    { header: '(+) y Blanco Cumple', key: 'cc_control_pos_blanco_ali_cumple', width: 16, group: 'CC' },
    { header: 'Ctrl Siembra Det.', key: 'cc_control_siembra_ali_detalle', width: 18, group: 'CC' },
    { header: 'Ctrl Siembra Anál.', key: 'cc_control_siembra_ali_analisis', width: 18, group: 'CC' },
    { header: 'Ctrl Siembra Cumple', key: 'cc_control_siembra_ali_cumple', width: 16, group: 'CC' },
    // CC2
    { header: 'CA2 Pesado T°', key: 'cc2_pesado_temp', width: 12, group: 'CC2' },
    { header: 'CA2 Pesado UFC', key: 'cc2_pesado_ufc', width: 12, group: 'CC2' },
    { header: 'CA2 Siembra', key: 'cc2_siembra', width: 16, group: 'CC2' },
    { header: 'Inicio', key: 'cc2_hora_inicio', width: 10, group: 'CC2' },
    { header: 'Término', key: 'cc2_hora_termino', width: 10, group: 'CC2' },
    { header: 'T°', key: 'cc2_temp', width: 8, group: 'CC2' },
    { header: 'E.coli UFC', key: 'cc2_ecoli_ufc', width: 12, group: 'CC2' },
    { header: 'Blanco UFC', key: 'cc2_blanco_ufc', width: 12, group: 'CC2' },
    // MIC
    { header: 'Desfavorable SI', key: 'mic_desfavorable_si', width: 14, group: 'MIC' },
    { header: 'Desfavorable NO', key: 'mic_desfavorable_no', width: 14, group: 'MIC' },
    { header: 'Tabla/Página', key: 'mic_tabla_pagina', width: 14, group: 'MIC' },
    { header: 'Límite', key: 'mic_limite', width: 12, group: 'MIC' },
    { header: 'Fecha Entrega', key: 'mic_fecha_entrega', width: 14, group: 'MIC' },
    { header: 'Hora Entrega', key: 'mic_hora_entrega', width: 12, group: 'MIC' },
    // Muestrario
    { header: 'Rep 1', key: 'muestrario_muestra_rep_1', width: 10, group: 'Muestrario' },
    { header: 'Rep 2', key: 'muestrario_muestra_rep_2', width: 10, group: 'Muestrario' },
    { header: 'Dil 1', key: 'muestrario_dil_1', width: 10, group: 'Muestrario' },
    { header: 'Dil 2', key: 'muestrario_dil_2', width: 10, group: 'Muestrario' },
    { header: 'c1-1', key: 'muestrario_c1_1', width: 10, group: 'Muestrario' },
    { header: 'c1-2', key: 'muestrario_c1_2', width: 10, group: 'Muestrario' },
    { header: 'c2-1', key: 'muestrario_c2_1', width: 10, group: 'Muestrario' },
    { header: 'c2-2', key: 'muestrario_c2_2', width: 10, group: 'Muestrario' },
    { header: 'ΣC1', key: 'muestrario_sumc_1', width: 10, group: 'Muestrario' },
    { header: 'ΣC2', key: 'muestrario_sumc_2', width: 10, group: 'Muestrario' },
    { header: 'd1', key: 'muestrario_d_1', width: 8, group: 'Muestrario' },
    { header: 'd2', key: 'muestrario_d_2', width: 8, group: 'Muestrario' },
    { header: 'n1-1', key: 'muestrario_n1_1', width: 8, group: 'Muestrario' },
    { header: 'n1-2', key: 'muestrario_n1_2', width: 8, group: 'Muestrario' },
    { header: 'n2-1', key: 'muestrario_n2_1', width: 8, group: 'Muestrario' },
    { header: 'n2-2', key: 'muestrario_n2_2', width: 8, group: 'Muestrario' },
    { header: 'x1', key: 'muestrario_x_1', width: 8, group: 'Muestrario' },
    { header: 'x2', key: 'muestrario_x_2', width: 8, group: 'Muestrario' },
    { header: 'Resultado RAM 1', key: 'muestrario_resultado_ram_1', width: 16, group: 'Muestrario' },
    { header: 'Resultado RAM 2', key: 'muestrario_resultado_ram_2', width: 16, group: 'Muestrario' },
    { header: 'Resultado RPES 1', key: 'muestrario_resultado_rpes_1', width: 16, group: 'Muestrario' },
    { header: 'Resultado RPES 2', key: 'muestrario_resultado_rpes_2', width: 16, group: 'Muestrario' },
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
      await require('./exportRAMController').addSheetForSample(wb, sample_id);
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
