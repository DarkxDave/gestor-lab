const ExcelJS = require('exceljs');
const ramModel = require('../models/formRAM');

// Reusable: add a RAM sheet for a given sample into an existing workbook
exports.addSheetForSample = async (wb, sample_id, data) => {
  const ws = wb.addWorksheet('RAM');
  if (!data) data = await ramModel.getBySampleId(sample_id);
  data = data || {};

  // Base styles/helpers
  ws.properties.defaultRowHeight = 18;
  const borderThin = { style: 'thin', color: { argb: 'FF000000' } };
  const setBorder = (r1, c1, r2, c2) => {
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const cell = ws.getCell(r, c);
        cell.border = { top: borderThin, left: borderThin, bottom: borderThin, right: borderThin };
      }
    }
  };
  const check = (v) => (v === true || v === 1 || v === '1') ? '√' : '';
  const yesNoMark = (val, want) => (val === null || val === undefined) ? '' : (String(val) === String(want) ? '√' : '');

  // Title + code
  ws.mergeCells('B1:T1');
  ws.getCell('B1').value = 'TRAZABILIDAD ANÁLISIS: ENUMERACIÓN DE AEROBIOS MESÓFILOS (NCh 2659.Of 2002)';
  ws.getCell('B1').alignment = { horizontal: 'center' };
  ws.getCell('B1').font = { bold: true, size: 18 };

  ws.mergeCells('B2:T2');
  ws.getCell('B2').value = 'R-PR-SVVM-M-4-11 / 15-02-23';
  ws.getCell('B2').alignment = { horizontal: 'center' };
  ws.getCell('B2').font = { bold: true, size: 12 };

  // ALI + sample id band
  ws.mergeCells('I3:K3');
  const idCell = ws.getCell('I3');
  idCell.value = `${sample_id}`;
  idCell.alignment = { horizontal: 'center', vertical: 'middle' };
  idCell.font = { bold: true, size: 14 };
  idCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
  ws.getRow(3).height = 24;
  setBorder(3, 9, 3, 11);
  const ali = ws.getCell('H3'); ali.value = 'ALI'; ali.alignment = { horizontal: 'center', vertical: 'middle' }; ali.font = { bold: true };
  setBorder(3, 8, 3, 8);

  // Column widths for main content area (B..K)
  ws.getColumn(2).width = 28; // B labels
  for (let c = 3; c <= 11; c++) ws.getColumn(c).width = 14;

  // Section: Fechas y análisis
  ws.mergeCells('B5:K5'); ws.getCell('B5').value = 'Fechas y análisis'; ws.getCell('B5').alignment = { horizontal: 'center' }; ws.getCell('B5').font = { bold: true };
  ws.mergeCells('B6:G6'); ws.getCell('B6').value = 'Inicio Incubación (Día/Mes/Hora/Analista):'; ws.getCell('B6').font = { bold: true };
  ws.mergeCells('H6:K6'); ws.getCell('H6').value = 'Término Análisis (Día/Mes/Hora/Analista):'; ws.getCell('H6').font = { bold: true };
  // Left (inicio)
  ws.getCell('B7').value = 'Fecha'; ws.mergeCells('C7:D7'); ws.getCell('C7').value = data.inicio_incubacion_fecha || '';
  ws.getCell('B8').value = 'Hora'; ws.mergeCells('C8:D8'); ws.getCell('C8').value = data.inicio_incubacion_hora || '';
  ws.getCell('B9').value = 'Analista'; ws.mergeCells('C9:G9'); ws.getCell('C9').value = data.inicio_incubacion_analista || '';
  // Right (termino)
  ws.getCell('H7').value = 'Fecha'; ws.mergeCells('I7:J7'); ws.getCell('I7').value = data.termino_analisis_fecha || '';
  ws.getCell('H8').value = 'Hora'; ws.mergeCells('I8:J8'); ws.getCell('I8').value = data.termino_analisis_hora || '';
  ws.getCell('H9').value = 'Analista'; ws.mergeCells('I9:K9'); ws.getCell('I9').value = data.termino_analisis_analista || '';
  setBorder(6, 2, 9, 11);

  // Section: Control Ambiental
  ws.mergeCells('B11:K11'); ws.getCell('B11').value = 'Control Ambiental'; ws.getCell('B11').alignment = { horizontal: 'center' }; ws.getCell('B11').font = { bold: true };
  ws.getCell('B12').value = 'Control Ambiental Pesado: T°:'; ws.getCell('C12').value = data.ca_pesado_temp || '';
  ws.getCell('D12').value = 'UFC:'; ws.getCell('E12').value = data.ca_pesado_ufc || '';
  ws.getCell('F12').value = 'Control ambiental Siembra:'; ws.mergeCells('G12:K12'); ws.getCell('G12').value = data.ca_siembra || '';
  ws.getCell('B13').value = 'Control de siembra E. Coli (UFC):'; ws.mergeCells('C13:D13'); ws.getCell('C13').value = data.ca_ecoli_ufc || '';
  ws.getCell('E13').value = 'Blanco (UFC):'; ws.mergeCells('F13:K13'); ws.getCell('F13').value = data.ca_blanco_ufc || '';
  setBorder(12, 2, 13, 11);

  // Section: Siembra
  ws.mergeCells('B15:K15'); ws.getCell('B15').value = 'Siembra'; ws.getCell('B15').alignment = { horizontal: 'center' }; ws.getCell('B15').font = { bold: true };
  ws.mergeCells('B16:J16'); ws.getCell('B16').value = 'Tiempo entre homogenizado y siembra < 15 minutos'; ws.getCell('K16').value = check(data.siembra_tiempo_ok);
  ws.getCell('B17').value = 'Minutos'; ws.mergeCells('C17:D17'); ws.getCell('C17').value = (data.siembra_tiempo_minutos ?? '') + '';
  ws.mergeCells('E17:H17'); ws.getCell('E17').value = 'N° de Muestra (10gr/90ml)'; ws.mergeCells('I17:K17'); ws.getCell('I17').value = data.siembra_n_muestra_10g_90ml || '';
  ws.mergeCells('E18:H18'); ws.getCell('E18').value = 'N° de Muestra (50gr/450ml)'; ws.mergeCells('I18:K18'); ws.getCell('I18').value = data.siembra_n_muestra_50g_450ml || '';
  setBorder(16, 2, 18, 11);

  // Section: Controles de Calidad
  ws.mergeCells('B20:K20'); ws.getCell('B20').value = 'Controles de Calidad'; ws.getCell('B20').alignment = { horizontal: 'center' }; ws.getCell('B20').font = { bold: true };
  // Header row
  ws.getCell('B21').value = 'Ítem'; ws.mergeCells('C21:E21'); ws.getCell('C21').value = 'Detalle (texto)'; ws.mergeCells('F21:G21'); ws.getCell('F21').value = 'Análisis (texto)'; ws.getCell('H21').value = 'Cumple'; ws.getCell('I21').value = 'No Cumple';
  // Rows
  const ccRows = [
    ['Duplicado en ALI', 'cc_duplicado_ali_detalle', 'cc_duplicado_ali_analisis', 'cc_duplicado_ali_cumple'],
    ['Control (+) y Blanco en ALI', 'cc_control_pos_blanco_ali_detalle', 'cc_control_pos_blanco_ali_analisis', 'cc_control_pos_blanco_ali_cumple'],
    ['Control de Siembra en ALI', 'cc_control_siembra_ali_detalle', 'cc_control_siembra_ali_analisis', 'cc_control_siembra_ali_cumple'],
  ];
  let row = 22;
  ccRows.forEach(r => {
    const [label, detKey, anaKey, cumpleKey] = r;
    ws.getCell(`B${row}`).value = label;
    ws.mergeCells(`C${row}:E${row}`); ws.getCell(`C${row}`).value = data[detKey] || '';
    ws.mergeCells(`F${row}:G${row}`); ws.getCell(`F${row}`).value = data[anaKey] || '';
    ws.getCell(`H${row}`).value = yesNoMark(data[cumpleKey], 1);
    ws.getCell(`I${row}`).value = yesNoMark(data[cumpleKey], 0);
    row++;
  });
  setBorder(21, 2, row - 1, 11);

  // Section: Control de Calidad 2
  ws.mergeCells('B25:K25'); ws.getCell('B25').value = 'Control de Calidad 2'; ws.getCell('B25').alignment = { horizontal: 'center' }; ws.getCell('B25').font = { bold: true };
  ws.getCell('B26').value = 'Control Ambiental Pesado: T°:'; ws.getCell('C26').value = data.cc2_pesado_temp || '';
  ws.getCell('D26').value = 'UFC:'; ws.getCell('E26').value = data.cc2_pesado_ufc || '';
  ws.getCell('F26').value = 'Control ambiental Siembra:'; ws.mergeCells('G26:K26'); ws.getCell('G26').value = data.cc2_siembra || '';
  ws.getCell('B27').value = 'Hora inicio:'; ws.getCell('C27').value = data.cc2_hora_inicio || '';
  ws.getCell('D27').value = 'Hora término:'; ws.getCell('E27').value = data.cc2_hora_termino || '';
  ws.getCell('F27').value = 'T°:'; ws.getCell('G27').value = data.cc2_temp || '';
  ws.getCell('B28').value = 'Control de siembra E. Coli (UFC):'; ws.mergeCells('C28:D28'); ws.getCell('C28').value = data.cc2_ecoli_ufc || '';
  ws.getCell('E28').value = 'Blanco (UFC):'; ws.mergeCells('F28:K28'); ws.getCell('F28').value = data.cc2_blanco_ufc || '';
  setBorder(26, 2, 28, 11);

  // Section: MIC (Manual de Inocuidad...) Parte II Sección III Cap IV pto 1 y 2
  ws.mergeCells('B30:K30'); ws.getCell('B30').value = 'Manual de Inocuidad y Certificación Parte II Sección III Cap IV pto 1 y 2'; ws.getCell('B30').alignment = { horizontal: 'center' }; ws.getCell('B30').font = { bold: true };
  ws.getCell('B31').value = 'Desfavorable:'; ws.getCell('C31').value = 'SI'; ws.getCell('D31').value = check(data.mic_desfavorable_si); ws.getCell('E31').value = 'NO'; ws.getCell('F31').value = check(data.mic_desfavorable_no);
  ws.getCell('B32').value = 'Tabla/Página:'; ws.mergeCells('C32:D32'); ws.getCell('C32').value = data.mic_tabla_pagina || '';
  ws.getCell('E32').value = 'Límite:'; ws.mergeCells('F32:K32'); ws.getCell('F32').value = data.mic_limite || '';
  ws.getCell('B33').value = 'Fecha y hora de entrega:'; ws.getCell('C33').value = data.mic_fecha_entrega || ''; ws.getCell('D33').value = data.mic_hora_entrega || '';
  setBorder(31, 2, 33, 11);

  // Section: Muestrario
  ws.mergeCells('B35:K35'); ws.getCell('B35').value = 'Muestrario'; ws.getCell('B35').alignment = { horizontal: 'center' }; ws.getCell('B35').font = { bold: true };
  // Table header
  ws.getCell('B36').value = '';
  ws.mergeCells('C36:D36'); ws.getCell('C36').value = 'N° de Muestra:';
  ws.mergeCells('E36:F36'); ws.getCell('E36').value = 'Duplicado:';
  // Row: N°
  ws.getCell('B37').value = 'N°';
  ws.getCell('C37').value = sample_id; ws.getCell('D37').value = data.muestrario_muestra_rep_1 || '';
  ws.getCell('E37').value = sample_id; ws.getCell('F37').value = data.muestrario_muestra_rep_2 || '';
  // Row: Dil
  ws.getCell('B38').value = 'Dil.:';
  ws.mergeCells('C38:D38'); ws.getCell('C38').value = data.muestrario_dil_1 || '';
  ws.mergeCells('E38:F38'); ws.getCell('E38').value = data.muestrario_dil_2 || '';
  // c1
  ws.getCell('B39').value = 'N° de colonias (c):';
  ws.mergeCells('C39:D39'); ws.getCell('C39').value = data.muestrario_c1_1 || '';
  ws.mergeCells('E39:F39'); ws.getCell('E39').value = data.muestrario_c1_2 || '';
  // c2
  ws.getCell('B40').value = 'N° de colonias (c):';
  ws.mergeCells('C40:D40'); ws.getCell('C40').value = data.muestrario_c2_1 || '';
  ws.mergeCells('E40:F40'); ws.getCell('E40').value = data.muestrario_c2_2 || '';
  // sumC
  ws.getCell('B41').value = '∑C (*)';
  ws.mergeCells('C41:D41'); ws.getCell('C41').value = data.muestrario_sumc_1 || '';
  ws.mergeCells('E41:F41'); ws.getCell('E41').value = data.muestrario_sumc_2 || '';
  // d
  ws.getCell('B42').value = 'd (*)';
  ws.mergeCells('C42:D42'); ws.getCell('C42').value = data.muestrario_d_1 || '';
  ws.mergeCells('E42:F42'); ws.getCell('E42').value = data.muestrario_d_2 || '';
  // n1
  ws.getCell('B43').value = 'n1 (*)';
  ws.mergeCells('C43:D43'); ws.getCell('C43').value = data.muestrario_n1_1 || '';
  ws.mergeCells('E43:F43'); ws.getCell('E43').value = data.muestrario_n1_2 || '';
  // n2
  ws.getCell('B44').value = 'n2 (*)';
  ws.mergeCells('C44:D44'); ws.getCell('C44').value = data.muestrario_n2_1 || '';
  ws.mergeCells('E44:F44'); ws.getCell('E44').value = data.muestrario_n2_2 || '';
  // x
  ws.getCell('B45').value = 'x (*)';
  ws.mergeCells('C45:D45'); ws.getCell('C45').value = data.muestrario_x_1 || '';
  ws.mergeCells('E45:F45'); ws.getCell('E45').value = data.muestrario_x_2 || '';
  // Resultado RAM
  ws.getCell('B46').value = 'Resultado RAM/Analista Lectura:';
  ws.getCell('C46').value = data.muestrario_resultado_ram_1 || '';
  ws.getCell('D46').value = 'UFC/g';
  ws.getCell('E46').value = data.muestrario_resultado_ram_2 || '';
  ws.getCell('F46').value = 'UFC/g';
  // Resultado RPES
  ws.getCell('B47').value = 'Resultado RPES/Analista Lectura:';
  ws.getCell('C47').value = data.muestrario_resultado_rpes_1 || '';
  ws.getCell('D47').value = 'UFC/g';
  ws.getCell('E47').value = data.muestrario_resultado_rpes_2 || '';
  ws.getCell('F47').value = 'UFC/g';
  setBorder(36, 2, 47, 11);

  // Section: Notas y Observaciones
  ws.mergeCells('B49:K49'); ws.getCell('B49').value = 'Notas y Observaciones'; ws.getCell('B49').alignment = { horizontal: 'center' }; ws.getCell('B49').font = { bold: true };
  ws.getCell('B50').value = 'Notas'; ws.mergeCells('C50:K52'); ws.getCell('C50').value = data.notes || ''; ws.getCell('C50').alignment = { vertical: 'top', wrapText: true };
  ws.getCell('B53').value = 'Observaciones'; ws.mergeCells('C53:K56'); ws.getCell('C53').value = data.observaciones || ''; ws.getCell('C53').alignment = { vertical: 'top', wrapText: true };
  setBorder(50, 2, 56, 11);

  // Signature area
  ws.mergeCells('B58:E58'); ws.getCell('B58').value = ''; // line placeholder
  ws.getCell('B59').value = 'FIRMA COORDINADOR DE ÁREA'; ws.getCell('B59').alignment = { horizontal: 'center' }; ws.getCell('B59').font = { bold: true };

  return ws;
};

exports.exportRAMForm = async (req, res, next) => {
  try {
    const sample_id = req.query.sample_id;
    if (!sample_id) return res.status(400).send('Parámetro sample_id requerido. Ej: /export/ram-form?sample_id=XYZ');

    const data = await ramModel.getBySampleId(sample_id);
    const wb = new ExcelJS.Workbook();
    await exports.addSheetForSample(wb, sample_id, data);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="RAM_${sample_id}.xlsx"`);
    await wb.xlsx.write(res);
    res.end();
  } catch (err) { next(err); }
};
