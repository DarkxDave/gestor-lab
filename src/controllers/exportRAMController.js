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
  const splitPair = (val) => {
    const s = (val == null) ? '' : String(val);
    const parts = s.split(/\s*[^0-9A-Za-z.,-]+\s*/).filter(p => p.length > 0);
    return [parts[0] || '', parts[1] || ''];
  };
  const toNumber = (v) => {
    if (v === null || v === undefined) return null;
    const s = String(v).trim().replace(',', '.');
    if (s === '') return null;
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  };

  // Title + code
  ws.mergeCells('B1:K1');
  ws.getCell('B1').value = 'TRAZABILIDAD ANÁLISIS: ENUMERACIÓN DE AEROBIOS MESÓFILOS (NCh 2659.Of 2002)';
  ws.getCell('B1').alignment = { horizontal: 'center' };
  ws.getCell('B1').font = { bold: true, size: 18 };

  ws.mergeCells('B2:K2');
  ws.getCell('B2').value = 'R-PR-SVVM-M-4-11 / 15-02-23';
  ws.getCell('B2').alignment = { horizontal: 'center' };
  ws.getCell('B2').font = { bold: true, size: 12 };

  // ALI + sample id band (moved to center-left: D3 for ALI, E3:G3 for ID)
  ws.mergeCells('E3:G3');
  const idCell = ws.getCell('E3');
  idCell.value = `${sample_id}`;
  idCell.alignment = { horizontal: 'center', vertical: 'middle' };
  idCell.font = { bold: true, size: 14 };
  idCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
  ws.getRow(3).height = 24;
  setBorder(3, 5, 3, 7);
  const ali = ws.getCell('D3'); ali.value = 'ALI'; ali.alignment = { horizontal: 'center', vertical: 'middle' }; ali.font = { bold: true };
  setBorder(3, 4, 3, 4);

  // Column widths for main content area (B..K)
  ws.getColumn(2).width = 28; // B labels
  for (let c = 1; c <= 26; c++) ws.getColumn(c).width = 14;

  // Section: Fechas y análisis
  // Achicar L, M, N, O como la columna D
  ws.getColumn(12).width = 4; // L
  ws.getColumn(13).width = 4; // M
  ws.getColumn(14).width = 4; // N
  ws.getColumn(15).width = 4; // O
  // Ajuste de ancho para columnas auxiliares S, V, Z
  ws.getColumn(19).width = 18; // S
  ws.getColumn(22).width = 18; // V
  ws.getColumn(26).width = 18; // Z
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
  // Clear specific borders requested: E7, F7, G7, K7 and E8, F8, G8, K8
  ['E7','F7','G7','K7','E8','F8','G8','K8'].forEach(addr => { ws.getCell(addr).border = undefined; });

  // Section: Set de Análisis (reemplaza Control Ambiental y CC2)
  ws.mergeCells('B11:K11'); ws.getCell('B11').value = 'Set de Análisis'; ws.getCell('B11').alignment = { horizontal: 'center' }; ws.getCell('B11').font = { bold: true };
  // Fila 1
  ws.getCell('B12').value = 'Control Ambiental Pesado: T°:'; ws.getCell('C12').value = data.cc2_pesado_temp || '';
  ws.getCell('D12').value = 'UFC:'; ws.getCell('E12').value = data.cc2_pesado_ufc || '';
  ws.getCell('F12').value = 'Control ambiental Siembra:'; ws.mergeCells('G12:K12'); ws.getCell('G12').value = data.cc2_siembra || '';
  // Fila 2
  ws.getCell('B13').value = 'Hora inicio:'; ws.getCell('C13').value = data.cc2_hora_inicio || '';
  ws.getCell('D13').value = 'Hora término:'; ws.getCell('E13').value = data.cc2_hora_termino || '';
  ws.getCell('F13').value = 'T°:'; ws.getCell('G13').value = data.cc2_temp || '';
  // Fila 3
  ws.getCell('B14').value = 'Control de siembra E. Coli (UFC):'; ws.mergeCells('C14:D14'); ws.getCell('C14').value = data.cc2_ecoli_ufc || '';
  ws.getCell('E14').value = 'Blanco (UFC):'; ws.mergeCells('F14:K14'); ws.getCell('F14').value = data.cc2_blanco_ufc || '';
  setBorder(12, 2, 14, 11);
  // Clear extra boxes on row 13 at I, J, K
  ['I13','J13','K13'].forEach(addr => { const cell = ws.getCell(addr); cell.value = ''; cell.border = undefined; });

  // Section: Siembra
  ws.mergeCells('B16:K16'); ws.getCell('B16').value = 'Siembra'; ws.getCell('B16').alignment = { horizontal: 'center' }; ws.getCell('B16').font = { bold: true };
  ws.mergeCells('B17:J17'); ws.getCell('B17').value = 'Tiempo entre homogenizado y siembra < 15 minutos'; ws.getCell('K17').value = check(data.siembra_tiempo_ok);
  ws.mergeCells('E18:H18'); ws.getCell('E18').value = 'N° de Muestra (10gr/90ml)'; ws.mergeCells('I18:K18'); ws.getCell('I18').value = data.siembra_n_muestra_10g_90ml || '';
  ws.mergeCells('E19:H19'); ws.getCell('E19').value = 'N° de Muestra (50gr/450ml)'; ws.mergeCells('I19:K19'); ws.getCell('I19').value = data.siembra_n_muestra_50g_450ml || '';
  setBorder(17, 2, 19, 11);

  // Section: MIC (Manual de Inocuidad...) Parte II Sección III Cap IV pto 1 y 2
  ws.mergeCells('B21:K21'); ws.getCell('B21').value = 'Manual de Inocuidad y Certificación Parte II Sección III Cap IV pto 1 y 2'; ws.getCell('B21').alignment = { horizontal: 'center' }; ws.getCell('B21').font = { bold: true };
  ws.getCell('B22').value = 'Desfavorable:'; ws.getCell('C22').value = 'SI'; ws.getCell('D22').value = check(data.mic_desfavorable_si); ws.getCell('E22').value = 'NO'; ws.getCell('F22').value = check(data.mic_desfavorable_no);
  ws.getCell('B23').value = 'Tabla/Página:'; ws.mergeCells('C23:D23'); ws.getCell('C23').value = data.mic_tabla_pagina || '';
  ws.getCell('E23').value = 'Límite:'; ws.mergeCells('F23:K23'); ws.getCell('F23').value = data.mic_limite || '';
  ws.getCell('B24').value = 'Fecha y hora de entrega:'; ws.getCell('C24').value = data.mic_fecha_entrega || ''; ws.getCell('D24').value = data.mic_hora_entrega || '';
  setBorder(22, 2, 24, 11);

  // Section: Datos (reemplazo con layout de ram_provisorio a partir de B27)
  ws.mergeCells('B26:K26'); ws.getCell('B26').value = 'Datos'; ws.getCell('B26').alignment = { horizontal: 'center' }; ws.getCell('B26').font = { bold: true };
  // Entradas base
  ws.mergeCells('B27:E27'); ws.getCell('B27').value = 'Suspension Inicial 1/:';
  ws.getCell('F27').value = data.datos_suspension_inicial_den || '';
  ws.mergeCells('B28:E28'); ws.getCell('B28').value = 'Volumen/Petri dish [mL]:';
  ws.getCell('F28').value = data.datos_volumen_petri_ml || '';

  // Encabezados de tabla (alineados con ram_provisorio)
  ws.mergeCells('B31:B32'); ws.getCell('B31').value = 'Dilusión'; ws.getCell('B31').alignment = { horizontal: 'center', vertical: 'middle' }; ws.getCell('B31').font = { bold: true };
  ws.mergeCells('C31:E31'); ws.getCell('C31').value = 'Numero de colonias'; ws.getCell('C31').alignment = { horizontal: 'center' }; ws.getCell('C31').font = { bold: true };
  ws.mergeCells('F31:G31'); ws.getCell('F31').value = ' Colonias por Confirmar'; ws.getCell('F31').alignment = { horizontal: 'center' }; ws.getCell('F31').font = { bold: true };
  ws.mergeCells('H31:I31'); ws.getCell('H31').value = 'Colonias Confirmadas'; ws.getCell('H31').alignment = { horizontal: 'center' }; ws.getCell('H31').font = { bold: true };
  ws.mergeCells('J31:K31'); ws.getCell('J31').value = 'Numero final de colonias'; ws.getCell('J31').alignment = { horizontal: 'center' }; ws.getCell('J31').font = { bold: true };
  // Sub-encabezados
  ws.getCell('C32').value = 'Lamina A';
  ws.mergeCells('D32:E32'); ws.getCell('D32').value = 'Lamina B';
  ws.getCell('F32').value = 'Lamina A'; ws.getCell('G32').value = 'Lamina B';
  ws.getCell('H32').value = 'Lamina A'; ws.getCell('I32').value = 'Lamina B';
  ws.getCell('J32').value = 'Lamina A'; ws.getCell('K32').value = 'Lamina B';
  ;['C32','D32','F32','G32','H32','I32','J32','K32'].forEach(addr => { ws.getCell(addr).alignment = { horizontal: 'center' }; });

  // Filas de captura (5) y merge D..E
  for (let r = 33; r <= 37; r++) {
    ws.mergeCells(`D${r}:E${r}`);
  }
  // Reglas de dilución en columna B (B33 entrada; B34..B37 derivadas)
  ws.getCell('B34').value = { formula: 'IF(B33="","",B33-1)' };
  ws.getCell('B35').value = { formula: 'IF(B34="","",B34-1)' };
  ws.getCell('B36').value = { formula: 'IF(B35="","",B35-1)' };
  ws.getCell('B37').value = { formula: 'IF(B36="","",B36-1)' };
  setBorder(31, 2, 37, 11);

  // Numero final de colonias (Lámina A y B) para todas las diluciones (1..5) en RAM: filas 33..37
  // Lámina A → columna J; usa C (num colonias), F (por confirmar), H (confirmadas)
  ws.getCell('J33').value = { formula: 'IF(F33>C33,"ERROR",IF(AND(F33="",H33=""),"",IF(IF(ISBLANK(F33)=ISBLANK(H33),TRUE,FALSE),C33*IF(F33<H33,"ERROR",H33/IF(F33=0,1,F33)),"ERROR")))' };
  ws.getCell('J34').value = { formula: 'IF(F34>C34,"ERROR",IF(AND(F34="",H34=""),"",IF(IF(ISBLANK(F34)=ISBLANK(H34),TRUE,FALSE),C34*IF(F34<H34,"ERROR",H34/IF(F34=0,1,F34)),"ERROR")))' };
  ws.getCell('J35').value = { formula: 'IF(F35>C35,"ERROR",IF(AND(F35="",H35=""),"",IF(IF(ISBLANK(F35)=ISBLANK(H35),TRUE,FALSE),C35*IF(F35<H35,"ERROR",H35/IF(F35=0,1,F35)),"ERROR")))' };
  ws.getCell('J36').value = { formula: 'IF(F36>C36,"ERROR",IF(AND(F36="",H36=""),"",IF(IF(ISBLANK(F36)=ISBLANK(H36),TRUE,FALSE),C36*IF(F36<H36,"ERROR",H36/IF(F36=0,1,F36)),"ERROR")))' };
  ws.getCell('J37').value = { formula: 'IF(F37>C37,"ERROR",IF(AND(F37="",H37=""),"",IF(IF(ISBLANK(F37)=ISBLANK(H37),TRUE,FALSE),C37*IF(F37<H37,"ERROR",H37/IF(F37=0,1,F37)),"ERROR")))' };
  // Lámina B → columna K; usa D (num colonias, celdas D..E unidas), G (por confirmar), I (confirmadas)
  ws.getCell('K33').value = { formula: 'IF(G33>D33,"ERROR",IF(AND(G33="",I33=""),"",IF(IF(ISBLANK(G33)=ISBLANK(I33),TRUE,FALSE),D33*IF(G33<I33,"ERROR",I33/IF(G33=0,1,G33)),"ERROR")))' };
  ws.getCell('K34').value = { formula: 'IF(G34>D34,"ERROR",IF(AND(G34="",I34=""),"",IF(IF(ISBLANK(G34)=ISBLANK(I34),TRUE,FALSE),D34*IF(G34<I34,"ERROR",I34/IF(G34=0,1,G34)),"ERROR")))' };
  ws.getCell('K35').value = { formula: 'IF(G35>D35,"ERROR",IF(AND(G35="",I35=""),"",IF(IF(ISBLANK(G35)=ISBLANK(I35),TRUE,FALSE),D35*IF(G35<I35,"ERROR",I35/IF(G35=0,1,G35)),"ERROR")))' };
  ws.getCell('K36').value = { formula: 'IF(G36>D36,"ERROR",IF(AND(G36="",I36=""),"",IF(IF(ISBLANK(G36)=ISBLANK(I36),TRUE,FALSE),D36*IF(G36<I36,"ERROR",I36/IF(G36=0,1,G36)),"ERROR")))' };
  ws.getCell('K37').value = { formula: 'IF(G37>D37,"ERROR",IF(AND(G37="",I37=""),"",IF(IF(ISBLANK(G37)=ISBLANK(I37),TRUE,FALSE),D37*IF(G37<I37,"ERROR",I37/IF(G37=0,1,G37)),"ERROR")))' };

  // Fila de resultado (notación científica)
  ws.mergeCells('B38:D38'); ws.getCell('B38').value = 'Numero de  CFU/g o mL:';
  ws.getCell('E38').value = '';
  ws.getCell('E38').font = { bold: true }; ws.getCell('E38').alignment = { horizontal: 'center' }; ws.getCell('E38').numFmt = '0.0E+00';
  ws.getCell('E38').border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };

  // Preguntas y respuestas (5 paralelas + 1 entre diluciones) y Observaciones reubicadas
  // Paralelo 1ª dilución
  ws.mergeCells('B40:I40'); ws.getCell('B40').value = '¿Es aceptable la diferencia de recuentos entre placas usadas en paralelo en la primera dilución?'; ws.getCell('B40').alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
  const a40 = ws.getCell('J40'); a40.font = { bold: true }; a40.alignment = { horizontal: 'center', vertical: 'middle' }; a40.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; a40.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
  // Paralelo 2ª dilución
  ws.mergeCells('B42:I42'); ws.getCell('B42').value = '¿Es aceptable la diferencia de recuentos entre placas usadas en paralelo en la segunda dilución?'; ws.getCell('B42').alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
  const a42 = ws.getCell('J42'); a42.font = { bold: true }; a42.alignment = { horizontal: 'center', vertical: 'middle' }; a42.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; a42.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
  // Paralelo 3ª dilución
  ws.mergeCells('B44:I44'); ws.getCell('B44').value = '¿Es aceptable la diferencia de recuentos entre placas usadas en paralelo en la tercera dilución?'; ws.getCell('B44').alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
  const a44p = ws.getCell('J44'); a44p.font = { bold: true }; a44p.alignment = { horizontal: 'center', vertical: 'middle' }; a44p.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; a44p.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
  // Paralelo 4ª dilución
  ws.mergeCells('B46:I46'); ws.getCell('B46').value = '¿Es aceptable la diferencia de recuentos entre placas usadas en paralelo en la cuarta dilución?'; ws.getCell('B46').alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
  const a46p = ws.getCell('J46'); a46p.font = { bold: true }; a46p.alignment = { horizontal: 'center', vertical: 'middle' }; a46p.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; a46p.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
  // Paralelo 5ª dilución
  ws.mergeCells('B48:I48'); ws.getCell('B48').value = '¿Es aceptable la diferencia de recuentos entre placas usadas en paralelo en la quinta dilución?'; ws.getCell('B48').alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
  const a48p = ws.getCell('J48'); a48p.font = { bold: true }; a48p.alignment = { horizontal: 'center', vertical: 'middle' }; a48p.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; a48p.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
  // Entre diluciones (global)
  ws.mergeCells('B50:G50'); ws.getCell('B50').value = '¿Es aceptable la diferencia de recuentos entre diluciones?'; ws.getCell('B50').alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
  ws.mergeCells('H50:I50'); const a50 = ws.getCell('H50'); a50.font = { bold: true }; a50.alignment = { horizontal: 'center', vertical: 'middle' }; a50.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; a50.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };

  // Observaciones (ahora en B52)
  ws.getCell('B52').value = 'Observaciones'; ws.mergeCells('C52:K56'); ws.getCell('C52').value = data.observaciones || ''; ws.getCell('C52').alignment = { vertical: 'top', wrapText: true };
  setBorder(52, 3, 56, 11);

  // Signature area (desplazada más abajo)
  ws.mergeCells('B58:E58'); ws.getCell('B58').value = '';
  ws.getCell('B59').value = 'FIRMA COORDINADOR DE ÁREA'; ws.getCell('B59').alignment = { horizontal: 'center' }; ws.getCell('B59').font = { bold: true };

  // Ajustes de anchos auxiliares para esta hoja
  ws.getColumn(16).width = 30; // P
  ws.getColumn(22).width = 18; // V
  ws.getColumn(26).width = 18; // Z

  // Rótulos auxiliares en columnas P, S y V (desplazados para iniciar en fila 25)
  // Columna P (equivalentes a P2..P21 en provisorio → P25..P44 en RAM)
  ws.getCell('P25').value = 'Dilución de la suspensión o volumen dado';
  ws.getCell('P26').value = 'Número de colonias dado';
  ws.getCell('P27').value = 'Número de colonias confirmadas';
  ws.getCell('P29').value = 'Duplicados 1ª dilución';
  ws.getCell('P30').value = 'MIN';
  ws.getCell('P31').value = 'MAX';
  ws.getCell('P32').value = 'P';
  ws.getCell('P33').value = 'CHISQ';
  ws.getCell('P37').value = 'Suma C si conf.';
  ws.getCell('P38').value = 'Suma C final';
  ws.getCell('P39').value = 'd';
  ws.getCell('P40').value = 'N° placas 1ª dilución';
  ws.getCell('P41').value = 'N° placas 2ª dilución';
  ws.getCell('P42').value = 'Volumen considerando diluciones';
  ws.getCell('P43').value = 'N';
  ws.getCell('P44').value = 'Probabilidad';
  ;['P25','P26','P27','P29','P30','P31','P32','P33','P37','P38','P39','P40','P41','P42','P43','P44'].forEach(addr => { ws.getCell(addr).font = { bold: true }; });

  // Columna S (S2..S10 → S25..S33)
  ws.getCell('S25').value = 'Initial susp.& Dilution';
  ws.getCell('S26').value = 'Only number is C10';
  ws.getCell('S27').value = 'All to confirm confirmed';
  ws.getCell('S29').value = 'Duplicates 2nd dilution';
  ws.getCell('S30').value = 'MIN';
  ws.getCell('S31').value = 'MAX';
  ws.getCell('S32').value = 'P';
  ws.getCell('S33').value = 'CHISQ';
  ;['S25','S26','S27','S29','S30','S31','S32','S33'].forEach(addr => { ws.getCell(addr).font = { bold: true }; });

  // Columna V (V2..V10 → V25..V33)
  ws.getCell('V25').value = 'ROUND';
  ws.getCell('V26').value = 'Check decimal C10';
  ws.getCell('V29').value = 'Dilution';
  ws.getCell('V30').value = 'SUM 1st';
  ws.getCell('V31').value = 'SUM 2nd';
  ws.getCell('V32').value = 'Pmin';
  ws.getCell('V33').value = 'CHISQmin';
  ;['V25','V26','V29','V30','V31','V32','V33'].forEach(addr => { ws.getCell(addr).font = { bold: true }; });

  // Celdas auxiliares Q/T/W/X con fórmulas desplazadas
  ws.getCell('E38').value = { formula: 'IF(T25="NO","Correct dilution factor",IF(Q25="NO","Error. Missingdata",IF(Q26="NO","",IF(T27="NO","",Q43))))' };
  // Aceptación paralelo 1ª..5ª y entre diluciones (J40,J42,J44,J46,J48,H50)
  ws.getCell('J40').value = { formula: 'IF(E38="","",IF(Q31=0,"NOT APPLICABLE",IF(Q40<2,"NOT APPLICABLE",IF(Q33>1-Q44,"YES","NO"))))' };
  ws.getCell('J42').value = { formula: 'IF(E38="","",IF(T31=0,"NOT APPLICABLE",IF(Q41<2,"NOT APPLICABLE",IF(T33>1-Q44,"YES","NO"))))' };
  ws.getCell('J44').value = { formula: 'IF(E38="","",IF(MAX(C35:E35)=0,"NOT APPLICABLE",IF(Q45<2,"NOT APPLICABLE",IF(T41>1-Q44,"YES","NO"))))' };
  ws.getCell('J46').value = { formula: 'IF(E38="","",IF(MAX(C36:E36)=0,"NOT APPLICABLE",IF(Q46<2,"NOT APPLICABLE",IF(T46>1-Q44,"YES","NO"))))' };
  ws.getCell('J48').value = { formula: 'IF(E38="","",IF(MAX(C37:E37)=0,"NOT APPLICABLE",IF(Q47<2,"NOT APPLICABLE",IF(T51>1-Q44,"YES","NO"))))' };
  ws.getCell('H50').value = { formula: 'IF(E38="","",IF(SUM(W30:W31)=0,"NOT APPLICABLE",IF(Q41=0,"NOT APPLICABLE",IF(W33>1-Q44,"YES","NO"))))' };
  // Q helpers
  ws.getCell('Q25').value = { formula: 'IF(OR(F27="",F28="",B33=""),"NO","YES")' };
  ws.getCell('Q26').value = { formula: 'IF(AND(C33="",D33="",C34="",D34=""),"NO","YES")' };
  ws.getCell('Q27').value = { formula: 'IF(AND(J33="",K33="",J34="",K34=""),"NO","YES")' };
  ws.getCell('Q30').value = { formula: 'MIN(C33:E33)' };
  ws.getCell('Q31').value = { formula: 'MAX(C33:E33)' };
  ws.getCell('Q32').value = { formula: '2*(Q30*LN(IF(Q30=0,1,Q30)/AVERAGE(Q30:Q31))+Q31*LN(Q31/AVERAGE(Q30:Q31)))' };
  ws.getCell('Q33').value = { formula: '1-CHISQ.DIST(Q32,1,TRUE)' };
  ws.getCell('Q37').value = { formula: 'IF(T27="N/A","",SUM(J33:K37))' };
  ws.getCell('Q38').value = { formula: 'IF(Q26="NO","",IF(MID(C33,1,1)=">",MID(C33,2,4),IF(Q37="",SUM(C33:E37),Q37)))' };
  ws.getCell('Q39').value = { formula: 'IF(F27=1,10^(B33),(1/F27)*(10^(B33+1)))' };
  ws.getCell('Q40').value = { formula: 'IF(Q27="NO",COUNTA(C33:E33),COUNTA(F33:G33))' };
  ws.getCell('Q41').value = { formula: 'IF(Q27="NO",COUNTA(C34:E34),COUNTA(F34:G34))' };
  // Placas contadas por dilución 3..5 y volumen ponderado extendido
  ws.getCell('Q45').value = { formula: 'IF(AND(J35="",K35=""),COUNTA(C35:E35),COUNTA(F35:G35))' };
  ws.getCell('Q46').value = { formula: 'IF(AND(J36="",K36=""),COUNTA(C36:E36),COUNTA(F36:G36))' };
  ws.getCell('Q47').value = { formula: 'IF(AND(J37="",K37=""),COUNTA(C37:E37),COUNTA(F37:G37))' };
  ws.getCell('Q42').value = { formula: 'F28*(Q40+(0.1*Q41)+(0.01*Q45)+(0.001*Q46)+(0.0001*Q47))' };
  ws.getCell('Q43').value = { formula: 'IF(Q38=0,1/(Q39*F28),Q38/(Q42*Q39))' };
  ws.getCell('Q44').value = 0.99;
  // T helpers
  ws.getCell('T25').value = { formula: 'IF(F27>1,IF(B33=0,"NO","YES"),"YES")' };
  ws.getCell('T26').value = { formula: 'IF(OR(ISNUMBER(C33),MID(C33,1,1)=">"),IF(X25=X24,"YES","decimal"),"NO")' };
  ws.getCell('T27').value = { formula: 'IF(Q27="NO","N/A",IF(OR(J33="ERROR",K33="ERROR",J34="ERROR",K34="ERROR"),"NO","YES"))' };
  ws.getCell('T30').value = { formula: 'MIN(C34:E34)' };
  ws.getCell('T31').value = { formula: 'MAX(C34:E34)' };
  ws.getCell('T32').value = { formula: '2*(T30*LN(IF(T30=0,1,T30)/AVERAGE(T30:T31))+T31*LN(T31/AVERAGE(T30:T31)))' };
  ws.getCell('T33').value = { formula: '1-CHISQ.DIST(T32,1,TRUE)' };
  // W helpers
  ws.getCell('W30').value = { formula: 'SUM(C33:E33)' };
  ws.getCell('W31').value = { formula: 'SUM(C34:E34)' };
  ws.getCell('W32').value = { formula: '2*(W30*LN(IF(W30=0,1,W30)/(10*(W30+W31)/11))+W31*LN(IF(W31=0,1,W31)/(1*(W30+W31)/11)))' };
  ws.getCell('W33').value = { formula: '1-CHISQ.DIST(W32,1,TRUE)' };
  // X helpers
  ws.getCell('X24').value = { formula: 'X26*1' };
  ws.getCell('X25').value = { formula: 'ROUND(X26,0)' };
  ws.getCell('X26').value = { formula: 'RIGHT(C33,2)' };

  // Bloques adicionales de duplicados para 3ra, 4ta y 5ta dilución (labels en S; cálculos en T)
  // 3ra dilución (fila de datos 35)
  ws.getCell('S37').value = '3ra dilusion';
  ws.getCell('S38').value = 'MIN';
  ws.getCell('S39').value = 'MAX';
  ws.getCell('S40').value = 'P';
  ws.getCell('S41').value = 'CHISQ';
  ;['S37','S38','S39','S40','S41'].forEach(addr => { ws.getCell(addr).font = { bold: true }; });
  ws.getCell('T38').value = { formula: 'MIN(C35:E35)' };
  ws.getCell('T39').value = { formula: 'MAX(C35:E35)' };
  ws.getCell('T40').value = { formula: '2*(T38*LN(IF(T38=0,1,T38)/AVERAGE(T38:T39))+T39*LN(T39/AVERAGE(T38:T39)))' };
  ws.getCell('T41').value = { formula: '1-CHISQ.DIST(T40,1,TRUE)' };

  // 4ta dilución (fila de datos 36)
  ws.getCell('S42').value = '4ta dilusion';
  ws.getCell('S43').value = 'MIN';
  ws.getCell('S44').value = 'MAX';
  ws.getCell('S45').value = 'P';
  ws.getCell('S46').value = 'CHISQ';
  ;['S42','S43','S44','S45','S46'].forEach(addr => { ws.getCell(addr).font = { bold: true }; });
  ws.getCell('T43').value = { formula: 'MIN(C36:E36)' };
  ws.getCell('T44').value = { formula: 'MAX(C36:E36)' };
  ws.getCell('T45').value = { formula: '2*(T43*LN(IF(T43=0,1,T43)/AVERAGE(T43:T44))+T44*LN(T44/AVERAGE(T43:T44)))' };
  ws.getCell('T46').value = { formula: '1-CHISQ.DIST(T45,1,TRUE)' };

  // 5ta dilución (fila de datos 37)
  ws.getCell('S47').value = '5ta dilusion';
  ws.getCell('S48').value = 'MIN';
  ws.getCell('S49').value = 'MAX';
  ws.getCell('S50').value = 'P';
  ws.getCell('S51').value = 'CHISQ';
  ;['S47','S48','S49','S50','S51'].forEach(addr => { ws.getCell(addr).font = { bold: true }; });
  ws.getCell('T48').value = { formula: 'MIN(C37:E37)' };
  ws.getCell('T49').value = { formula: 'MAX(C37:E37)' };
  ws.getCell('T50').value = { formula: '2*(T48*LN(IF(T48=0,1,T48)/AVERAGE(T48:T49))+T49*LN(T49/AVERAGE(T48:T49)))' };
  ws.getCell('T51').value = { formula: '1-CHISQ.DIST(T50,1,TRUE)' };

  // Aceptación (paralelo) para 3ra–5ta dilución en RAM (resultados tipo J40/J42)
  ws.getCell('R37').value = 'OK 3ra'; ws.getCell('R37').font = { bold: true };
  ws.getCell('R41').value = { formula: 'IF(E38="","",IF(MAX(C35:E35)=0,"NOT APPLICABLE",IF(COUNTA(C35:E35)<2,"NOT APPLICABLE",IF(T41>1-Q44,"YES","NO"))))' };
  ws.getCell('R42').value = 'OK 4ta'; ws.getCell('R42').font = { bold: true };
  ws.getCell('R46').value = { formula: 'IF(E38="","",IF(MAX(C36:E36)=0,"NOT APPLICABLE",IF(COUNTA(C36:E36)<2,"NOT APPLICABLE",IF(T46>1-Q44,"YES","NO"))))' };
  ws.getCell('R47').value = 'OK 5ta'; ws.getCell('R47').font = { bold: true };
  ws.getCell('R51').value = { formula: 'IF(E38="","",IF(MAX(C37:E37)=0,"NOT APPLICABLE",IF(COUNTA(C37:E37)<2,"NOT APPLICABLE",IF(T51>1-Q44,"YES","NO"))))' };

  return ws;
};

exports.exportRAMForm = async (req, res, next) => {
  try {
    const sample_id = req.query.sample_id;
    if (!sample_id) return res.status(400).send('Parámetro sample_id requerido. Ej: /export/ram-form?sample_id=XYZ');

    const data = await ramModel.getBySampleId(sample_id);
    const wb = new ExcelJS.Workbook();
    await exports.addSheetForSample(wb, sample_id, data);
  // Provisional sheet eliminado según requerimiento (solo hoja RAM principal)

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="RAM_${sample_id}.xlsx"`);
    await wb.xlsx.write(res);
    res.end();
  } catch (err) { next(err); }
};