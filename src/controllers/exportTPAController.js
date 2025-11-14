const ExcelJS = require('exceljs');
const tpaModel = require('../models/tpaFormModel');

function asCheck(v) {
  return (v === true || v === 1 || v === '1') ? '✓' : '';
}


// Reusable helper: añade una hoja TPA a un workbook dado
exports.addSheetForSample = async (wb, sample_id, rec) => {
  if (!rec) rec = await tpaModel.getBySampleId(sample_id);
  const ws = wb.addWorksheet('TPA');

  // Helpers
  const check = (v) => asCheck(v);
  const setBorder = (r1, c1, r2, c2) => {
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        ws.getCell(r, c).border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      }
    }
  };

  // ID band and ALI label
  ws.mergeCells('I3:K3');
  const idCell = ws.getCell('I3');
  idCell.value = `${sample_id}`;
  idCell.alignment = { horizontal: 'center', vertical: 'middle' };
  idCell.font = { bold: true, size: 14 };
  idCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
  ws.getRow(3).height = 24;
  setBorder(3, 9, 3, 11); // I3..K3
  const aliCell = ws.getCell('H3');
  aliCell.value = 'ALI';
  aliCell.alignment = { horizontal: 'center', vertical: 'middle' };
  aliCell.font = { bold: true };
  setBorder(3, 8, 3, 8);

  // Anchos de columnas
  ws.getColumn(2).width = 35; // B
  ws.getColumn(3).width = 18; // C
  ws.getColumn(4).width = 18; // D
  ws.getColumn(5).width = 18; // E
  ws.getColumn(6).width = 18; // F
  ws.getColumn(7).width = 8;  // G
  ws.getColumn(8).width = 40; // H

  // Almacenamiento + Observaciones
  ws.mergeCells('B5:G5');
  ws.getCell('B5').value = 'Lugar de almacenamiento de muestras:';
  ws.getCell('B5').alignment = { horizontal: 'center' };
  ws.getCell('B5').font = { bold: true, size: 12 };

  ws.mergeCells('H5:H9');
  ws.getCell('H5').value = 'Observaciones:';
  ws.getCell('H5').alignment = { vertical: 'top' };
  ws.getCell('H5').font = { bold: true };

  setBorder(5, 2, 9, 8); // B5..H9

  const items = [
    ['Frezeer 33-M', 'storage_freezer_33m'],
    ['Refrigerador 33-M', 'storage_refrigerador_33m'],
    ['Mesón siembra', 'storage_meson_siembra'],
    ['Gabinete sala Traspaso', 'storage_gabinete_traspaso'],
  ];
  ['B6', 'B7', 'B8', 'B9'].forEach((addr, i) => {
    ws.mergeCells(`${addr}:F${6 + i}`);
    ws.getCell(addr).value = items[i][0];
    ws.getCell(addr).alignment = { horizontal: 'center' };
    ws.getCell(`G${6 + i}`).value = check(rec && rec[items[i][1]]);
    ws.getCell(`G${6 + i}`).alignment = { horizontal: 'center' };
  });
  ws.getCell('H6').value = (rec && rec.observaciones) || '';
  ws.getCell('H6').alignment = { wrapText: true, vertical: 'top' };

  // Columnas adicionales
  ws.getColumn(9).width = 14;  // I
  ws.getColumn(10).width = 16; // J
  ws.getColumn(11).width = 18; // K
  ws.getColumn(12).width = 4;  // L
  ws.getColumn(13).width = 18; // M
  ws.getColumn(14).width = 18; // N
  ws.getColumn(15).width = 18; // O
  ws.getColumn(16).width = 8;  // P
  ws.getColumn(17).width = 3;  // Q
  ws.getColumn(18).width = 22; // R
  ws.getColumn(19).width = 22; // S
  ws.getColumn(20).width = 8;  // T

  // Título sección principal
  ws.mergeCells('B11:K11');
  ws.getCell('B11').value = 'MANIPULACIÓN DE MUESTRAS (1)';
  ws.getCell('B11').alignment = { horizontal: 'center' };
  ws.getCell('B11').font = { bold: true, size: 12 };

  // Encabezados de tabla (manipulación 1)
  ws.mergeCells('B12:B13'); ws.getCell('B12').value = 'Retiro de Muestra'; ws.getCell('B12').alignment = { horizontal: 'center', vertical: 'middle' }; ws.getCell('B12').font = { bold: true };
  ws.mergeCells('C12:C13'); ws.getCell('C12').value = 'Pesado'; ws.getCell('C12').alignment = { horizontal: 'center', vertical: 'middle' }; ws.getCell('C12').font = { bold: true };
  ws.mergeCells('D12:F12'); ws.getCell('D12').value = 'Clave material pesado (*)'; ws.getCell('D12').alignment = { horizontal: 'center' }; ws.getCell('D12').font = { bold: true };
  ws.mergeCells('G12:G13'); ws.getCell('G12').value = 'Responsable'; ws.getCell('G12').alignment = { horizontal: 'center', vertical: 'middle' }; ws.getCell('G12').font = { bold: true };
  ws.mergeCells('H12:H13'); ws.getCell('H12').value = 'Fecha'; ws.getCell('H12').alignment = { horizontal: 'center', vertical: 'middle' }; ws.getCell('H12').font = { bold: true };
  ws.mergeCells('I12:I13'); ws.getCell('I12').value = 'Hora de inicio'; ws.getCell('I12').alignment = { horizontal: 'center', vertical: 'middle' }; ws.getCell('I12').font = { bold: true };
  ws.mergeCells('J12:J13'); ws.getCell('J12').value = 'Hora de término/Inicio de almacenamiento'; ws.getCell('J12').alignment = { horizontal: 'center', vertical: 'middle' }; ws.getCell('J12').font = { bold: true };
  ws.mergeCells('K12:K13'); ws.getCell('K12').value = 'N° de muestra'; ws.getCell('K12').alignment = { horizontal: 'center', vertical: 'middle' }; ws.getCell('K12').font = { bold: true };

  ws.getCell('D13').value = 'Cuchara:'; ws.getCell('D13').alignment = { horizontal: 'center' };
  ws.getCell('E13').value = 'Pinzas:'; ws.getCell('E13').alignment = { horizontal: 'center' };
  ws.getCell('F13').value = 'Bisturí:'; ws.getCell('F13').alignment = { horizontal: 'center' };

  const filas = [
    { retiro: 'retiro_muestra_1', pesado: 'pesado_1', cm1: 'clave_material_1', cm2: 'clave_material_2', cm3: 'clave_material_3', resp: 'responsable_1', fecha: 'fecha_1', h1: 'hora_inicio_1', h2: 'hora_termino_1', n: 'n_muestra_1' },
    { retiro: 'retiro_muestra_2', pesado: 'pesado_2', cm1: 'clave_material_4', cm2: 'clave_material_5', cm3: 'clave_material_6', resp: 'responsable_2', fecha: 'fecha_2', h1: 'hora_inicio_2', h2: 'hora_termino_2', n: 'n_muestra_2' },
    { retiro: 'retiro_muestra_3', pesado: 'pesado_3', cm1: 'clave_material_7', cm2: 'clave_material_8', cm3: 'clave_material_9', resp: 'responsable_3', fecha: 'fecha_3', h1: 'hora_inicio_3', h2: 'hora_termino_3', n: 'n_muestra_3' },
  ];
  [14, 15, 16].forEach((r, i) => {
    const f = filas[i];
    ws.getCell(`B${r}`).value = check(rec && rec[f.retiro]);
    ws.getCell(`C${r}`).value = check(rec && rec[f.pesado]);
    ws.getCell(`D${r}`).value = rec && rec[f.cm1] || '';
    ws.getCell(`E${r}`).value = rec && rec[f.cm2] || '';
    ws.getCell(`F${r}`).value = rec && rec[f.cm3] || '';
    ws.getCell(`G${r}`).value = rec && rec[f.resp] || '';
    const fechaVal = rec && rec[f.fecha];
    ws.getCell(`H${r}`).value = (fechaVal instanceof Date) ? fechaVal.toISOString().slice(0, 10) : (fechaVal || '');
    ws.getCell(`I${r}`).value = rec && rec[f.h1] || '';
    ws.getCell(`J${r}`).value = rec && rec[f.h2] || '';
    ws.getCell(`K${r}`).value = rec && rec[f.n] || '';
  });
  setBorder(11, 2, 16, 11);

  // Equipos para Pesado (B18..E)
  ws.mergeCells('B18:E18'); ws.getCell('B18').value = 'Equipos para Pesado:'; ws.getCell('B18').alignment = { horizontal: 'center' }; ws.getCell('B18').font = { bold: true };
  const eqRows = [
    ['Balanza  74-M', 'equipo_balanza_74m'],
    ['Cámara flujo laminar 8-M', 'equipo_camara_8m'],
    ['Balanza  6-M', 'equipo_balanza_6m'],
    ['Mesón de  traspaso', 'equipo_meson_traspaso'],
    ['Balanza  99-M', 'equipo_balanza_99m'],
    ['Balanza  108-M', 'equipo_balanza_108m', true],
  ];
  eqRows.forEach((row, i) => {
    const rr = 19 + i;
    ws.mergeCells(`B${rr}:D${rr}`);
    ws.getCell(`B${rr}`).value = row[0];
    if (row[2]) ws.getCell(`B${rr}`).font = { bold: true };
    ws.getCell(`B${rr}`).alignment = { horizontal: 'center' };
    ws.getCell(`E${rr}`).value = check(rec && rec[row[1]]);
    ws.getCell(`E${rr}`).alignment = { horizontal: 'center' };
  });
  setBorder(18, 2, 24, 5); // B18..E24

  // Lugar de almacenamiento (G18..I)
  ws.mergeCells('G18:I18'); ws.getCell('G18').value = 'Lugar de almacenamiento:'; ws.getCell('G18').alignment = { horizontal: 'center' }; ws.getCell('G18').font = { bold: true };
  const locRows = [
    ['Frezeer 33-M', 'storage_freezer_33m'],
    ['Refrigerador 33-M', 'storage_refrigerador_33m'],
    ['Gabinete de Traspaso', 'storage_gabinete_traspaso'],
  ];
  locRows.forEach((row, i) => {
    const rr = 19 + i;
    ws.mergeCells(`G${rr}:H${rr}`);
    ws.getCell(`G${rr}`).value = row[0];
    ws.getCell(`G${rr}`).alignment = { horizontal: 'center' };
    ws.getCell(`I${rr}`).value = check(rec && rec[row[1]]);
    ws.getCell(`I${rr}`).alignment = { horizontal: 'center' };
  });
  setBorder(18, 7, 21, 9); // G18..I21

  // Puntas/pipetas desechables para pesado (*) (J18..L)
  const baseP = 18;
  ws.mergeCells(`J${baseP}:L${baseP}`);
  ws.getCell(`J${baseP}`).value = 'Puntas/pipetas desechables para pesado';
  ws.getCell(`J${baseP}`).alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getCell(`J${baseP}`).font = { bold: true };
  ws.getCell(`J${baseP}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };

  ws.mergeCells(`J${baseP + 1}:L${baseP + 1}`);
  ws.getCell(`J${baseP + 1}`).value = '(*)';
  ws.getCell(`J${baseP + 1}`).alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getCell(`J${baseP + 1}`).font = { bold: true };
  ws.getCell(`J${baseP + 1}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };

  ws.getCell(`J${baseP + 2}`).value = '1ml /clave:';
  ws.mergeCells(`K${baseP + 2}:L${baseP + 2}`);
  ws.getCell(`K${baseP + 2}`).value = (rec && rec.clave_1ml) || '';

  ws.getCell(`J${baseP + 3}`).value = '10 ml/clave:';
  ws.mergeCells(`K${baseP + 3}:L${baseP + 3}`);
  ws.getCell(`K${baseP + 3}`).value = (rec && rec.clave_10ml) || '';

  ws.mergeCells(`J${baseP + 4}:J${baseP + 6}`);
  ws.getCell(`J${baseP + 4}`).value = 'Otros:';
  ws.getCell(`J${baseP + 4}`).alignment = { vertical: 'middle' };
  ws.mergeCells(`K${baseP + 4}:L${baseP + 6}`);
  ws.getCell(`K${baseP + 4}`).value = (rec && rec.clave_otros) || '';
  ws.getCell(`K${baseP + 4}`).alignment = { vertical: 'top', wrapText: true };

  setBorder(baseP, 10, baseP + 6, 12);

  // Uso y limpieza de Micropipetas
  let r0 = baseP + 8;
  ws.mergeCells(`B${r0}:K${r0}`);
  ws.getCell(`B${r0}`).value = 'Uso y limpieza de Micropipetas';
  ws.getCell(`B${r0}`).alignment = { horizontal: 'center' };
  ws.getCell(`B${r0}`).font = { bold: true };

  // Sub-headers para 1 ml y 10 ml
  r0++;
  ws.mergeCells(`B${r0}:C${r0}`); ws.getCell(`B${r0}`).value = '1 ml'; ws.getCell(`B${r0}`).alignment = { horizontal: 'center' }; ws.getCell(`B${r0}`).font = { bold: true };
  ws.mergeCells(`E${r0}:F${r0}`); ws.getCell(`E${r0}`).value = '10 ml'; ws.getCell(`E${r0}`).alignment = { horizontal: 'center' }; ws.getCell(`E${r0}`).font = { bold: true };
  ws.mergeCells(`H${r0}:K${r0}`); ws.getCell(`H${r0}`).value = 'Claves Micropipetas'; ws.getCell(`H${r0}`).alignment = { horizontal: 'center' }; ws.getCell(`H${r0}`).font = { bold: true };

  const mpLeft = [
    ['22-M', 'micropipeta_22m'],
    ['23-M', 'micropipeta_23m'],
    ['72-M', 'micropipeta_72m'],
    ['98-M', 'micropipeta_98m'],
    ['100-M', 'micropipeta_100m'],
    ['102-M', 'micropipeta_102m'],
    ['106-M', 'micropipeta_106m'],
  ];
  const mpRight = [
    ['32-M', 'micropipeta_32m'],
    ['75-M', 'micropipeta_75m'],
    ['94-M', 'micropipeta_94m'],
    ['103-M', 'micropipeta_103m'],
  ];
  let maxRows = Math.max(mpLeft.length, mpRight.length);
  for (let i = 0; i < maxRows; i++) {
    const rr = r0 + 1 + i;
    if (mpLeft[i]) {
      ws.getCell(`B${rr}`).value = mpLeft[i][0];
      ws.getCell(`C${rr}`).value = check(rec && rec[mpLeft[i][1]]);
      ws.getCell(`C${rr}`).alignment = { horizontal: 'center' };
    }
    if (mpRight[i]) {
      ws.getCell(`E${rr}`).value = mpRight[i][0];
      ws.getCell(`F${rr}`).value = check(rec && rec[mpRight[i][1]]);
      ws.getCell(`F${rr}`).alignment = { horizontal: 'center' };
    }
  }
  const clavesStart = r0 + 1;
  ws.getCell(`H${clavesStart}`).value = 'Puntas/pipetas 1ml/clave'; ws.mergeCells(`I${clavesStart}:K${clavesStart}`); ws.getCell(`I${clavesStart}`).value = (rec && rec.clave_1ml) || '';
  ws.getCell(`H${clavesStart + 1}`).value = 'Puntas/pipetas 10ml/clave'; ws.mergeCells(`I${clavesStart + 1}:K${clavesStart + 1}`); ws.getCell(`I${clavesStart + 1}`).value = (rec && rec.clave_10ml) || '';
  ws.getCell(`H${clavesStart + 2}`).value = 'Otros'; ws.mergeCells(`I${clavesStart + 2}:K${clavesStart + 2}`); ws.getCell(`I${clavesStart + 2}`).value = (rec && rec.clave_otros) || '';
  setBorder(r0, 2, r0 + maxRows, 6);
  setBorder(r0, 8, r0 + 3, 11);

  // Limpieza
  let rL = r0 + maxRows + 2;
  ws.mergeCells(`B${rL}:K${rL}`); ws.getCell(`B${rL}`).value = 'Limpieza'; ws.getCell(`B${rL}`).alignment = { horizontal: 'center' }; ws.getCell(`B${rL}`).font = { bold: true };
  const limpRows = [
    ['Mesón', 'limpieza_meson'],
    ['Stomacher 12-M', 'limpieza_stomacher'],
    ['Cámara flujo laminar 8-M', 'limpieza_camara'],
    ['Balanza 74-M 108-M', 'limpieza_balanza'],
    ['Balanza 6-M 99-M', 'limpieza_balanza2'],
  ];
  let rr = rL + 1;
  limpRows.forEach(item => {
    ws.mergeCells(`B${rr}:E${rr}`); ws.getCell(`B${rr}`).value = item[0];
    ws.getCell(`F${rr}`).value = check(rec && rec[item[1]]); ws.getCell(`F${rr}`).alignment = { horizontal: 'center' };
    rr++;
  });
  ws.mergeCells(`B${rr}:E${rr}`); ws.getCell(`B${rr}`).value = 'Otros'; ws.mergeCells(`F${rr}:K${rr}`); ws.getCell(`F${rr}`).value = (rec && rec.limpieza_otros) || '';
  rr++;
  ws.mergeCells(`B${rr}:E${rr}`); ws.getCell(`B${rr}`).value = 'Desinfectante en aerosol'; ws.getCell(`F${rr}`).value = check(rec && rec.limpieza_aerosol); ws.getCell(`F${rr}`).alignment = { horizontal: 'center' };
  rr++;
  ws.mergeCells(`B${rr}:E${rr}`); ws.getCell(`B${rr}`).value = 'Observaciones limpieza'; ws.mergeCells(`F${rr}:K${rr + 2}`); ws.getCell(`F${rr}`).value = (rec && rec.observaciones_limpieza) || ''; ws.getCell(`F${rr}`).alignment = { vertical: 'top', wrapText: true };
  setBorder(rL + 1, 2, rr + 2, 11);

  // Material y equipos de Siembra
  let rS = rr + 4;
  ws.mergeCells(`B${rS}:K${rS}`); ws.getCell(`B${rS}`).value = 'Material y equipos de Siembra'; ws.getCell(`B${rS}`).alignment = { horizontal: 'center' }; ws.getCell(`B${rS}`).font = { bold: true };
  rS++;
  ws.getCell(`B${rS}`).value = 'Clave'; ws.mergeCells(`C${rS}:K${rS}`); ws.getCell(`C${rS}`).value = (rec && rec.clave_general) || '';
  rS++;
  const siembraRows = [
    ['Puntas 1mL:', 'clave_puntas_1ml', 'Baño 5-M:', 'bano_5m', true],
    ['Puntas 10 mL:', 'clave_puntas_10ml', 'Homogenizador 12-M:', 'homogenizador_12m', true],
    ['Placas estériles: 57cm2 / 150 mm', 'clave_placas', 'Cuenta colonias 9-M:', 'cuenta_colonias_9m', true],
    ['Asas Drigalsky:', 'clave_asas', 'Cuenta colonias 101-M:', 'cuenta_colonias_101m', true],
    ['Blender:', 'clave_blender', 'pHmetro 93-M:', 'phmetro_93m', true],
    ['Bolsas estériles:', 'clave_bolsas', 'Pipetas desechables:', 'pipetas_desechables', true],
    ['Probeta 250 ml / 100 ml', 'clave_probeta', 'Otro:', 'clave_otro', false],
  ];
  siembraRows.forEach(row => {
    const [l1, k1, l2, k2, isCheck2] = row;
    ws.mergeCells(`B${rS}:D${rS}`); ws.getCell(`B${rS}`).value = l1; ws.mergeCells(`E${rS}:G${rS}`); ws.getCell(`E${rS}`).value = (rec && rec[k1]) || '';
    ws.mergeCells(`H${rS}:J${rS}`); ws.getCell(`H${rS}`).value = l2;
    if (isCheck2) { ws.getCell(`K${rS}`).value = check(rec && rec[k2]); ws.getCell(`K${rS}`).alignment = { horizontal: 'center' }; }
    else { ws.getCell(`K${rS}`).value = (rec && rec[k2]) || ''; }
    rS++;
  });
  setBorder(rS - siembraRows.length, 2, rS - 1, 11);

  // Diluyentes
  let rD = rS + 1;
  ws.mergeCells(`B${rD}:K${rD}`); ws.getCell(`B${rD}`).value = 'Diluyentes'; ws.getCell(`B${rD}`).alignment = { horizontal: 'center' }; ws.getCell(`B${rD}`).font = { bold: true };
  rD++;
  const dilRows = [
    ['AP 0,1% 90 ml:', 'ap_90ml', 'AP 0,1% tubos ml:', 'ap_tubos_ml'],
    ['AP 0,1% 99 ml:', 'ap_99ml', 'SPS 225 ml:', 'sps_225ml'],
    ['AP 0,1% 450 ml:', 'ap_450ml', 'SPS Tubos:', 'sps_tubos'],
    ['AP 0,1% 225 ml:', 'ap_225ml', 'SPS sa 90 ml:', 'sps_sa_90ml'],
    ['AP 0,1% 500 ml:', 'ap_500ml', 'SPS sa tubos:', 'sps_sa_tubos'],
    ['PBS 450 ml:', 'pbs_450ml', 'Otro:', 'diluyente_otro'],
    ['Otros:', 'diluyente_otros1', '', ''],
  ];
  dilRows.forEach(row => {
    const [l1, k1, l2, k2] = row;
    ws.mergeCells(`B${rD}:D${rD}`); ws.getCell(`B${rD}`).value = l1; ws.mergeCells(`E${rD}:G${rD}`); ws.getCell(`E${rD}`).value = (rec && rec[k1]) || '';
    if (l2) {
      ws.mergeCells(`H${rD}:J${rD}`); ws.getCell(`H${rD}`).value = l2; ws.getCell(`K${rD}`).value = (rec && rec[k2]) || '';
    } else {
      ws.mergeCells(`H${rD}:K${rD}`); ws.getCell(`H${rD}`).value = (rec && rec[k1]) || '';
    }
    rD++;
  });
  setBorder(rD - dilRows.length, 2, rD - 1, 11);

  // Firma
  let rF = rD + 2;
  ws.mergeCells(`B${rF}:E${rF}`);
  ws.getCell(`B${rF + 1}`).value = 'FIRMA COORDINADOR DE ÁREA';
  ws.getCell(`B${rF + 1}`).alignment = { horizontal: 'center' };
  ws.getCell(`B${rF + 1}`).font = { bold: true };

  return ws;
};

exports.exportExcel = async (req, res, next) => {
  try {
    const rows = await tpaModel.listAll();
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('TPA');

    // Define el esquema de columnas con grupos y subgrupos
    const schema = [
      { header: 'sample_id', key: 'sample_id', width: 16, group: 'Identificación' },

      // Almacenamiento
      { header: 'Freezer 33-M', key: 'storage_freezer_33m', width: 14, group: 'Almacenamiento' },
      { header: 'Refrigerador 33-M', key: 'storage_refrigerador_33m', width: 16, group: 'Almacenamiento' },
      { header: 'Mesón siembra', key: 'storage_meson_siembra', width: 14, group: 'Almacenamiento' },
      { header: 'Gabinete Traspaso', key: 'storage_gabinete_traspaso', width: 16, group: 'Almacenamiento' },
      { header: 'Observaciones', key: 'observaciones', width: 30, group: 'Almacenamiento' },

      // Manipulación (1)
      { header: 'Retiro', key: 'retiro_muestra_1', width: 10, group: 'Manipulación (1)' },
      { header: 'Pesado', key: 'pesado_1', width: 10, group: 'Manipulación (1)' },
      { header: 'Cuchara', key: 'clave_material_1', width: 12, group: 'Manipulación (1)', subgroup: 'Clave material pesado' },
      { header: 'Pinzas', key: 'clave_material_2', width: 12, group: 'Manipulación (1)', subgroup: 'Clave material pesado' },
      { header: 'Bisturí', key: 'clave_material_3', width: 12, group: 'Manipulación (1)', subgroup: 'Clave material pesado' },
      { header: 'Responsable', key: 'responsable_1', width: 18, group: 'Manipulación (1)' },
      { header: 'Fecha', key: 'fecha_1', width: 12, group: 'Manipulación (1)' },
      { header: 'H.Inicio', key: 'hora_inicio_1', width: 12, group: 'Manipulación (1)' },
      { header: 'H.Término', key: 'hora_termino_1', width: 12, group: 'Manipulación (1)' },
      { header: 'Nº muestra', key: 'n_muestra_1', width: 14, group: 'Manipulación (1)' },

      // Manipulación (2)
      { header: 'Retiro', key: 'retiro_muestra_2', width: 10, group: 'Manipulación (2)' },
      { header: 'Pesado', key: 'pesado_2', width: 10, group: 'Manipulación (2)' },
      { header: 'Cuchara', key: 'clave_material_4', width: 12, group: 'Manipulación (2)', subgroup: 'Clave material pesado' },
      { header: 'Pinzas', key: 'clave_material_5', width: 12, group: 'Manipulación (2)', subgroup: 'Clave material pesado' },
      { header: 'Bisturí', key: 'clave_material_6', width: 12, group: 'Manipulación (2)', subgroup: 'Clave material pesado' },
      { header: 'Responsable', key: 'responsable_2', width: 18, group: 'Manipulación (2)' },
      { header: 'Fecha', key: 'fecha_2', width: 12, group: 'Manipulación (2)' },
      { header: 'H.Inicio', key: 'hora_inicio_2', width: 12, group: 'Manipulación (2)' },
      { header: 'H.Término', key: 'hora_termino_2', width: 12, group: 'Manipulación (2)' },
      { header: 'Nº muestra', key: 'n_muestra_2', width: 14, group: 'Manipulación (2)' },

      // Manipulación (3)
      { header: 'Retiro', key: 'retiro_muestra_3', width: 10, group: 'Manipulación (3)' },
      { header: 'Pesado', key: 'pesado_3', width: 10, group: 'Manipulación (3)' },
      { header: 'Cuchara', key: 'clave_material_7', width: 12, group: 'Manipulación (3)', subgroup: 'Clave material pesado' },
      { header: 'Pinzas', key: 'clave_material_8', width: 12, group: 'Manipulación (3)', subgroup: 'Clave material pesado' },
      { header: 'Bisturí', key: 'clave_material_9', width: 12, group: 'Manipulación (3)', subgroup: 'Clave material pesado' },
      { header: 'Responsable', key: 'responsable_3', width: 18, group: 'Manipulación (3)' },
      { header: 'Fecha', key: 'fecha_3', width: 12, group: 'Manipulación (3)' },
      { header: 'H.Inicio', key: 'hora_inicio_3', width: 12, group: 'Manipulación (3)' },
      { header: 'H.Término', key: 'hora_termino_3', width: 12, group: 'Manipulación (3)' },
      { header: 'Nº muestra', key: 'n_muestra_3', width: 14, group: 'Manipulación (3)' },

      // Equipos/Lugares
      { header: 'Balanza 74-M', key: 'equipo_balanza_74m', width: 14, group: 'Equipos/Lugares' },
      { header: 'Cámara 8-M', key: 'equipo_camara_8m', width: 12, group: 'Equipos/Lugares' },
      { header: 'Balanza 6-M', key: 'equipo_balanza_6m', width: 12, group: 'Equipos/Lugares' },
      { header: 'Mesón Traspaso', key: 'equipo_meson_traspaso', width: 14, group: 'Equipos/Lugares' },
      { header: 'Balanza 99-M', key: 'equipo_balanza_99m', width: 12, group: 'Equipos/Lugares' },
      { header: 'Balanza 108-M', key: 'equipo_balanza_108m', width: 14, group: 'Equipos/Lugares' },
      { header: 'Freezer 33-M', key: 'equipo_freezer_33m', width: 12, group: 'Equipos/Lugares' },
      { header: 'Refrigerador 33-M', key: 'equipo_refrigerador_33m', width: 16, group: 'Equipos/Lugares' },
      { header: 'Gabinete Traspaso', key: 'equipo_gabinete_traspaso', width: 16, group: 'Equipos/Lugares' },

      // Micropipetas (subgrupos 1 ml / 10 ml)
      { header: '22-M', key: 'micropipeta_22m', width: 8, group: 'Micropipetas', subgroup: '1 ml' },
      { header: '23-M', key: 'micropipeta_23m', width: 8, group: 'Micropipetas', subgroup: '1 ml' },
      { header: '72-M', key: 'micropipeta_72m', width: 8, group: 'Micropipetas', subgroup: '1 ml' },
      { header: '98-M', key: 'micropipeta_98m', width: 8, group: 'Micropipetas', subgroup: '1 ml' },
      { header: '100-M', key: 'micropipeta_100m', width: 9, group: 'Micropipetas', subgroup: '1 ml' },
      { header: '102-M', key: 'micropipeta_102m', width: 9, group: 'Micropipetas', subgroup: '1 ml' },
      { header: '106-M', key: 'micropipeta_106m', width: 9, group: 'Micropipetas', subgroup: '1 ml' },
      { header: '32-M', key: 'micropipeta_32m', width: 8, group: 'Micropipetas', subgroup: '10 ml' },
      { header: '75-M', key: 'micropipeta_75m', width: 8, group: 'Micropipetas', subgroup: '10 ml' },
      { header: '94-M', key: 'micropipeta_94m', width: 8, group: 'Micropipetas', subgroup: '10 ml' },
      { header: '103-M', key: 'micropipeta_103m', width: 9, group: 'Micropipetas', subgroup: '10 ml' },
      { header: 'Clave 1ml', key: 'clave_1ml', width: 12, group: 'Micropipetas' },
      { header: 'Clave 10ml', key: 'clave_10ml', width: 12, group: 'Micropipetas' },
      { header: 'Clave otros', key: 'clave_otros', width: 18, group: 'Micropipetas' },

      // Limpieza
      { header: 'Limp. Mesón', key: 'limpieza_meson', width: 12, group: 'Limpieza' },
      { header: 'Limp. Stomacher', key: 'limpieza_stomacher', width: 14, group: 'Limpieza' },
      { header: 'Limp. Cámara', key: 'limpieza_camara', width: 12, group: 'Limpieza' },
      { header: 'Limp. Balanza', key: 'limpieza_balanza', width: 12, group: 'Limpieza' },
      { header: 'Limp. Balanza2', key: 'limpieza_balanza2', width: 13, group: 'Limpieza' },
      { header: 'Limp. Otros', key: 'limpieza_otros', width: 18, group: 'Limpieza' },
      { header: 'Aerosol', key: 'limpieza_aerosol', width: 10, group: 'Limpieza' },
      { header: 'Obs. Limpieza', key: 'observaciones_limpieza', width: 25, group: 'Limpieza' },

      // Siembra
      { header: 'Clave General', key: 'clave_general', width: 16, group: 'Siembra' },
      { header: 'Puntas 1mL', key: 'clave_puntas_1ml', width: 14, group: 'Siembra' },
      { header: 'Baño 5-M', key: 'bano_5m', width: 10, group: 'Siembra' },
      { header: 'Puntas 10mL', key: 'clave_puntas_10ml', width: 14, group: 'Siembra' },
      { header: 'Homog. 12-M', key: 'homogenizador_12m', width: 12, group: 'Siembra' },
      { header: 'Placas', key: 'clave_placas', width: 12, group: 'Siembra' },
      { header: 'Cuenta 9-M', key: 'cuenta_colonias_9m', width: 12, group: 'Siembra' },
      { header: 'Asas', key: 'clave_asas', width: 10, group: 'Siembra' },
      { header: 'Cuenta 101-M', key: 'cuenta_colonias_101m', width: 12, group: 'Siembra' },
      { header: 'Blender', key: 'clave_blender', width: 12, group: 'Siembra' },
      { header: 'pHmetro 93-M', key: 'phmetro_93m', width: 14, group: 'Siembra' },
      { header: 'Bolsas', key: 'clave_bolsas', width: 10, group: 'Siembra' },
      { header: 'Pipetas desech.', key: 'pipetas_desechables', width: 14, group: 'Siembra' },
      { header: 'Probeta', key: 'clave_probeta', width: 12, group: 'Siembra' },
      { header: 'Otro', key: 'clave_otro', width: 16, group: 'Siembra' },

      // Diluyentes
      { header: 'AP 0,1% 90 ml', key: 'ap_90ml', width: 14, group: 'Diluyentes' },
      { header: 'AP tubos ml', key: 'ap_tubos_ml', width: 12, group: 'Diluyentes' },
      { header: 'AP 0,1% 99 ml', key: 'ap_99ml', width: 14, group: 'Diluyentes' },
      { header: 'SPS 225 ml', key: 'sps_225ml', width: 12, group: 'Diluyentes' },
      { header: 'AP 0,1% 450 ml', key: 'ap_450ml', width: 16, group: 'Diluyentes' },
      { header: 'SPS Tubos', key: 'sps_tubos', width: 12, group: 'Diluyentes' },
      { header: 'AP 0,1% 225 ml', key: 'ap_225ml', width: 16, group: 'Diluyentes' },
      { header: 'SPS sa 90 ml', key: 'sps_sa_90ml', width: 14, group: 'Diluyentes' },
      { header: 'AP 0,1% 500 ml', key: 'ap_500ml', width: 16, group: 'Diluyentes' },
      { header: 'SPS sa tubos', key: 'sps_sa_tubos', width: 14, group: 'Diluyentes' },
      { header: 'PBS 450 ml', key: 'pbs_450ml', width: 12, group: 'Diluyentes' },
      { header: 'Diluyente otro', key: 'diluyente_otro', width: 16, group: 'Diluyentes' },
      { header: 'Diluyentes otros', key: 'diluyente_otros1', width: 16, group: 'Diluyentes' },

    ];

    // Configuracion de columnas
    ws.columns = schema.map(c => ({ key: c.key, width: c.width || 12 }));

    // Titulo principal (fila 1, combinada)
    const totalCols = schema.length;
    ws.mergeCells(1, 1, 1, totalCols);
    const titleCell = ws.getCell(1, 1);
    titleCell.value = 'Formulario de Trazabilidad (TPA) - Exportación';
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    titleCell.font = { bold: true, size: 14 };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F0FE' } };

    // Fila 2: Encabezados de grupo (combinados donde aplique)
    const groups = [];
    const groupOrder = [];
    schema.forEach((c, idx) => {
      const g = c.group || '';
      if (!groups[g]) { groups[g] = { start: idx + 1, end: idx + 1 }; groupOrder.push(g); }
      else groups[g].end = idx + 1;
    });
    groupOrder.forEach(g => {
      const { start, end } = groups[g];
      ws.mergeCells(2, start, 2, end);
      const cell = ws.getCell(2, start);
      cell.value = g || '';
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F6FA' } };
      // Dibuja bordes gruesos en los límites de los grupos
      ws.getCell(2, start).border = { left: { style: 'thick', color: { argb: 'FF9EB5D5' } } };
    });

    // Fila 3: Encabezados de subgrupo (combinados donde aplique)
    // Construye rangos de subgrupos
    const subgroupRanges = {};
    schema.forEach((c, idx) => {
      if (c.subgroup) {
        const key = `${c.group}__${c.subgroup}`;
        if (!subgroupRanges[key]) subgroupRanges[key] = { start: idx + 1, end: idx + 1, label: c.subgroup };
        else subgroupRanges[key].end = idx + 1;
      }
    });
    // Inicializa fila 3 con celdas vacías
    for (let i = 1; i <= totalCols; i++) ws.getCell(3, i).value = '';
    Object.values(subgroupRanges).forEach(r => {
      ws.mergeCells(3, r.start, 3, r.end);
      const cell = ws.getCell(3, r.start);
      cell.value = r.label;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF5FB' } };
    });

    // Fila 4: Encabezados de columna individuales
    ws.getRow(4).values = [, ...schema.map(c => c.header)];
    ws.getRow(4).font = { bold: true };
    ws.getRow(4).alignment = { horizontal: 'center' };

    // Bordes finos alrededor de encabezados (filas 1-4)
    [1, 2, 3, 4].forEach(rn => {
      const row = ws.getRow(rn);
      for (let i = 1; i <= totalCols; i++) {
        const cell = ws.getCell(rn, i);
        cell.border = cell.border || {};
        cell.border.top = cell.border.top || { style: 'thin', color: { argb: 'FFBFCAD9' } };
        cell.border.bottom = { style: 'thin', color: { argb: 'FFBFCAD9' } };
      }
    });

    // Las filas de datos comienzan en la fila 5
    rows.forEach(r => {
      // Ordenado usando el esquema; forzar booleanos/0-1 a marcas de verificación para columnas apropiadas
      const ordered = schema.map(c => {
        const v = r[c.key];
        // Heurística: tratar 0/1 numéricos y booleanos como valores de casilla de verificación para claves booleanas típicas
        const isBoolishKey = (
          c.key.startsWith('storage_') || c.key.startsWith('retiro_') || c.key.startsWith('pesado_') ||
          c.key.startsWith('equipo_') || c.key.startsWith('micropipeta_') || c.key.startsWith('limpieza_') ||
          ['bano_5m', 'homogenizador_12m', 'cuenta_colonias_9m', 'cuenta_colonias_101m', 'phmetro_93m', 'pipetas_desechables', 'b_approved', 'b_qc_pass']
            .includes(c.key)
        );
        if (isBoolishKey) return asCheck(v);
        if (['fecha_1', 'fecha_2', 'fecha_3'].includes(c.key) && v instanceof Date) return v.toISOString().slice(0, 10);
        return v ?? '';
      });
      const newRow = ws.addRow(ordered);
      newRow.alignment = { vertical: 'middle' };
    });

    // Inmovilizar filas de encabezado
    ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 4 }];

    // Opcional: bandas de cebra para filas de datos
    const firstDataRow = 5;
    for (let r = firstDataRow; r <= ws.rowCount; r++) {
      if ((r - firstDataRow) % 2 === 0) {
        for (let c = 1; c <= schema.length; c++) {
          ws.getCell(r, c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDFEFF' } };
        }
      }
    }

    // Setea bordes gruesos de grupo en fila de encabezados (fila 4)
    groupOrder.forEach(g => {
      const start = groups[g].start;
      const end = groups[g].end;
      // Bordes gruesos en los límites de grupo para la fila 4 (encabezados de columna)
      ws.getCell(4, start).border = {
        left: { style: 'thick', color: { argb: 'FF9EB5D5' } },
        top: { style: 'thin', color: { argb: 'FFBFCAD9' } },
        bottom: { style: 'thin', color: { argb: 'FFBFCAD9' } },
      };
      ws.getCell(4, end).border = {
        right: { style: 'thick', color: { argb: 'FF9EB5D5' } },
        top: { style: 'thin', color: { argb: 'FFBFCAD9' } },
        bottom: { style: 'thin', color: { argb: 'FFBFCAD9' } },
      };
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="tpa_muestras.xlsx"');
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    next(err);
  }
};

// Añade la hoja TPA
exports.exportTPAForm = async (req, res, next) => {
  try {
    const sample_id = req.query.sample_id;
    if (!sample_id) return res.status(400).send('Parámetro sample_id requerido. Ej: /export/tpa-form?sample_id=XYZ');
    const rec = await tpaModel.getBySampleId(sample_id);
    if (!rec) return res.status(404).send('Muestra no encontrada');

    const wb = new ExcelJS.Workbook();
  
  await exports.addSheetForSample(wb, sample_id, rec);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="TPA_${sample_id}.xlsx"`);
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    next(err);
  }
};
