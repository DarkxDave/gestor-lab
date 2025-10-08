const ExcelJS = require('exceljs');
const formAModel = require('../models/formA');

function asCheck(v) {
  return (v === true || v === 1 || v === '1') ? '✓' : '';
}

exports.exportExcel = async (req, res, next) => {
  try {
    const rows = await formAModel.listJoinFormB();
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('TPA');

    // Define schema with groups and optional subgroups to mirror on-page sections
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

      // Form B
      { header: 'B Notas', key: 'b_notes', width: 24, group: 'Formulario B' },
      { header: 'B Aprobado', key: 'b_approved', width: 12, group: 'Formulario B' },
      { header: 'B QC OK', key: 'b_qc_pass', width: 10, group: 'Formulario B' },
    ];

    // Setup worksheet columns
    ws.columns = schema.map(c => ({ key: c.key, width: c.width || 12 }));

    // Top title row
    const totalCols = schema.length;
    ws.mergeCells(1, 1, 1, totalCols);
    const titleCell = ws.getCell(1, 1);
    titleCell.value = 'Formulario de Trazabilidad (TPA) - Exportación';
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    titleCell.font = { bold: true, size: 14 };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F0FE' } };

    // Row 2: Group headings (merged)
    let col = 1;
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
      // Draw a thick left border at group boundaries for header rows
      ws.getCell(2, start).border = { left: { style: 'thick', color: { argb: 'FF9EB5D5' } } };
    });

    // Row 3: Subgroup headings (merged where applicable)
    // Build subgroup ranges within each group
    const subgroupRanges = {};
    schema.forEach((c, idx) => {
      if (c.subgroup) {
        const key = `${c.group}__${c.subgroup}`;
        if (!subgroupRanges[key]) subgroupRanges[key] = { start: idx + 1, end: idx + 1, label: c.subgroup };
        else subgroupRanges[key].end = idx + 1;
      }
    });
    // Initialize row 3 with empty cells
    for (let i = 1; i <= totalCols; i++) ws.getCell(3, i).value = '';
    Object.values(subgroupRanges).forEach(r => {
      ws.mergeCells(3, r.start, 3, r.end);
      const cell = ws.getCell(3, r.start);
      cell.value = r.label;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF5FB' } };
    });

    // Row 4: Column headers
    ws.getRow(4).values = [ , ...schema.map(c => c.header) ];
    ws.getRow(4).font = { bold: true };
    ws.getRow(4).alignment = { horizontal: 'center' };

    // Thin borders around header area
    [1,2,3,4].forEach(rn => {
      const row = ws.getRow(rn);
      for (let i = 1; i <= totalCols; i++) {
        const cell = ws.getCell(rn, i);
        cell.border = cell.border || {};
        cell.border.top = cell.border.top || { style: 'thin', color: { argb: 'FFBFCAD9' } };
        cell.border.bottom = { style: 'thin', color: { argb: 'FFBFCAD9' } };
      }
    });

    // Data rows start at row 5
    rows.forEach(r => {
      // Ordered using schema; coerce booleans/0-1 to checkmarks for appropriate columns
      const ordered = schema.map(c => {
        const v = r[c.key];
        // Heuristic: treat numeric 0/1 and booleans as checkbox values for typical boolean keys
        const isBoolishKey = (
          c.key.startsWith('storage_') || c.key.startsWith('retiro_') || c.key.startsWith('pesado_') ||
          c.key.startsWith('equipo_') || c.key.startsWith('micropipeta_') || c.key.startsWith('limpieza_') ||
          ['bano_5m','homogenizador_12m','cuenta_colonias_9m','cuenta_colonias_101m','phmetro_93m','pipetas_desechables','b_approved','b_qc_pass']
            .includes(c.key)
        );
        if (isBoolishKey) return asCheck(v);
        if (['fecha_1','fecha_2','fecha_3'].includes(c.key) && v instanceof Date) return v.toISOString().slice(0,10);
        return v ?? '';
      });
      const newRow = ws.addRow(ordered);
      newRow.alignment = { vertical: 'middle' };
    });

    // Freeze panes (keep sample_id and headers visible)
    ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 4 }];

    // Optional: zebra stripes for data rows
    const firstDataRow = 5;
    for (let r = firstDataRow; r <= ws.rowCount; r++) {
      if ((r - firstDataRow) % 2 === 0) {
        for (let c = 1; c <= schema.length; c++) {
          ws.getCell(r, c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDFEFF' } };
        }
      }
    }

    // Set borders at group boundaries for column header row to visually separate sections
    let running = 0;
    groupOrder.forEach(g => {
      const start = groups[g].start;
      const end = groups[g].end;
      // Thick left border at start, thick right border at end for row 4 (column headers)
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

exports.exportTPAForm = async (req, res, next) => {
  try {
    const sample_id = req.query.sample_id;
    if (!sample_id) return res.status(400).send('Parámetro sample_id requerido. Ej: /export/tpa-form?sample_id=XYZ');

    const rows = await require('../models/formA').listJoinFormB();
    const rec = rows.find(r => r.sample_id === sample_id);
    if (!rec) return res.status(404).send('Muestra no encontrada');

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('TPA');

  // Config base
    ws.properties.defaultRowHeight = 18;
    const borderThin = { style: 'thin', color: { argb: 'FF000000' } };

    // Helper
    const setBorder = (range) => {
      const [r1,c1,r2,c2] = range;
      for (let r=r1; r<=r2; r++) {
        for (let c=c1; c<=c2; c++) {
          const cell = ws.getCell(r,c);
          cell.border = {
            top: borderThin,
            left: borderThin,
            bottom: borderThin,
            right: borderThin,
          };
        }
      }
    };
    const merge = (r1,c1,r2,c2,val,opts={}) => {
      ws.mergeCells(r1,c1,r2,c2);
      const cell = ws.getCell(r1,c1);
      if (val !== undefined) cell.value = val;
      if (opts.align) cell.alignment = opts.align;
      if (opts.font) cell.font = opts.font;
      return cell;
    };
    // Preferimos el símbolo "√" como en tu ejemplo
    const check = (v) => (v===true || v===1 || v==='1') ? '√' : '';
  // ==============================
  // Encabezado superior (filas 1 y 2)
  // ==============================
  ws.mergeCells('B1:T1');
  ws.getCell('B1').value = 'TRAZABILIDAD Y ANÁLISIS';
  ws.getCell('B1').alignment = { horizontal: 'center' };
  ws.getCell('B1').font = { bold: true, size: 18 };

  ws.mergeCells('B2:T2');
  ws.getCell('B2').value = 'R-INS-MM-M-1-15 /23-08-23';
  ws.getCell('B2').alignment = { horizontal: 'center' };
  ws.getCell('B2').font = { bold: true, size: 12 };

  // Fila 3: ID de la muestra en la columna 3 (C)
  const sampleIdOnly = req.query.sample_id || (rec && rec.sample_id) || '';
  // Limpiamos C3 si lo estuviera usando antes
  ws.getCell('C3').value = '';
  // ID destacado debajo de la columna I
  ws.mergeCells('I3:K3');
  const idCell = ws.getCell('I3');
  idCell.value = `${sampleIdOnly}`;
  idCell.alignment = { horizontal: 'center', vertical: 'middle' };
  idCell.font = { bold: true, size: 14 };
  idCell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFF2F2F2'} };
  ws.getRow(3).height = 24;
  setBorder([3,9,3,11]); // I3..K3

  // Etiqueta ALI a la izquierda del ID (H3)
  const aliCell = ws.getCell('H3');
  aliCell.value = 'ALI';
  aliCell.alignment = { horizontal: 'center', vertical: 'middle' };
  aliCell.font = { bold: true };
  setBorder([3,8,3,8]); // H3


    // Ancho de columnas para cuadro de Almacenamiento (similar screenshot)
    // A vacía, B título, C-G tabla de almacenamiento, H observaciones
  ws.getColumn(2).width = 35; // B
  ws.getColumn(3).width = 18; // C
  ws.getColumn(4).width = 18; // D
  ws.getColumn(5).width = 18; // E
  ws.getColumn(6).width = 18; // F
  ws.getColumn(7).width = 8;  // G
  ws.getColumn(8).width = 40; // H

    // Bloque: Almacenamiento + Observaciones
    ws.mergeCells('B5:G5');
    ws.getCell('B5').value = 'Lugar de almacenamiento de muestras:';
    ws.getCell('B5').alignment = { horizontal: 'center' };
    ws.getCell('B5').font = { bold: true, size: 12 };

    ws.mergeCells('H5:H9');
    ws.getCell('H5').value = 'Observaciones:';
    ws.getCell('H5').alignment = { vertical: 'top' };
    ws.getCell('H5').font = { bold: true };

    setBorder([5,2,9,8]); // B5..H9

    const items = [
      ['Frezeer 33-M','storage_freezer_33m'],
      ['Refrigerador 33-M','storage_refrigerador_33m'],
      ['Mesón siembra','storage_meson_siembra'],
      ['Gabinete sala Traspaso','storage_gabinete_traspaso'],
    ];
    ['B6','B7','B8','B9'].forEach((addr, i) => {
      ws.mergeCells(`${addr}:F${6+i}`);
      ws.getCell(addr).value = items[i][0];
      ws.getCell(addr).alignment = { horizontal: 'center' };
      ws.getCell(`G${6+i}`).value = check(rec[items[i][1]]);
      ws.getCell(`G${6+i}`).alignment = { horizontal: 'center' };
    });

    ws.getCell('H6').value = rec.observaciones || '';
    ws.getCell('H6').alignment = { wrapText: true, vertical: 'top' };

    // ==============================
    // Sección: MANIPULACIÓN DE MUESTRAS (1)
    // ==============================
    // Columnas principales B..K; paneles derechos M..P y R..T
  ws.getColumn(9).width = 14;  // I opcional
  ws.getColumn(10).width = 16; // J
  ws.getColumn(11).width = 18; // K
  ws.getColumn(12).width = 4;  // L separador
  ws.getColumn(13).width = 18; // M
  ws.getColumn(14).width = 18; // N
  ws.getColumn(15).width = 18; // O
  ws.getColumn(16).width = 8;  // P marca
  ws.getColumn(17).width = 3;  // Q separador
  ws.getColumn(18).width = 22; // R
  ws.getColumn(19).width = 22; // S
  ws.getColumn(20).width = 8;  // T marca

    // Título
    ws.mergeCells('B11:K11');
    ws.getCell('B11').value = 'MANIPULACIÓN DE MUESTRAS (1)';
    ws.getCell('B11').alignment = { horizontal: 'center' };
    ws.getCell('B11').font = { bold: true, size: 12 };

    // Encabezados
    ws.mergeCells('B12:B13'); ws.getCell('B12').value = 'Retiro de Muestra'; ws.getCell('B12').alignment = { horizontal:'center', vertical:'middle' }; ws.getCell('B12').font = { bold:true };
    ws.mergeCells('C12:C13'); ws.getCell('C12').value = 'Pesado'; ws.getCell('C12').alignment = { horizontal:'center', vertical:'middle' }; ws.getCell('C12').font = { bold:true };
    ws.mergeCells('D12:F12'); ws.getCell('D12').value = 'Clave material pesado (*)'; ws.getCell('D12').alignment = { horizontal:'center' }; ws.getCell('D12').font = { bold:true };
    ws.mergeCells('G12:G13'); ws.getCell('G12').value = 'Responsable'; ws.getCell('G12').alignment = { horizontal:'center', vertical:'middle' }; ws.getCell('G12').font = { bold:true };
    ws.mergeCells('H12:H13'); ws.getCell('H12').value = 'Fecha'; ws.getCell('H12').alignment = { horizontal:'center', vertical:'middle' }; ws.getCell('H12').font = { bold:true };
    ws.mergeCells('I12:I13'); ws.getCell('I12').value = 'Hora de inicio'; ws.getCell('I12').alignment = { horizontal:'center', vertical:'middle' }; ws.getCell('I12').font = { bold:true };
    ws.mergeCells('J12:J13'); ws.getCell('J12').value = 'Hora de término/Inicio de almacenamiento'; ws.getCell('J12').alignment = { horizontal:'center', vertical:'middle' }; ws.getCell('J12').font = { bold:true };
    ws.mergeCells('K12:K13'); ws.getCell('K12').value = 'N° de muestra'; ws.getCell('K12').alignment = { horizontal:'center', vertical:'middle' }; ws.getCell('K12').font = { bold:true };

    ws.getCell('D13').value = 'Cuchara:'; ws.getCell('D13').alignment = { horizontal:'center' };
    ws.getCell('E13').value = 'Pinzas:';  ws.getCell('E13').alignment = { horizontal:'center' };
    ws.getCell('F13').value = 'Bisturí:'; ws.getCell('F13').alignment = { horizontal:'center' };

    // Filas de datos (14..16)
    const filas = [
      { retiro:'retiro_muestra_1', pesado:'pesado_1', cm1:'clave_material_1', cm2:'clave_material_2', cm3:'clave_material_3', resp:'responsable_1', fecha:'fecha_1', h1:'hora_inicio_1', h2:'hora_termino_1', n:'n_muestra_1' },
      { retiro:'retiro_muestra_2', pesado:'pesado_2', cm1:'clave_material_4', cm2:'clave_material_5', cm3:'clave_material_6', resp:'responsable_2', fecha:'fecha_2', h1:'hora_inicio_2', h2:'hora_termino_2', n:'n_muestra_2' },
      { retiro:'retiro_muestra_3', pesado:'pesado_3', cm1:'clave_material_7', cm2:'clave_material_8', cm3:'clave_material_9', resp:'responsable_3', fecha:'fecha_3', h1:'hora_inicio_3', h2:'hora_termino_3', n:'n_muestra_3' },
    ];
    [14,15,16].forEach((r, i) => {
      const f = filas[i];
      ws.getCell(`B${r}`).value = check(rec[f.retiro]);
      ws.getCell(`C${r}`).value = check(rec[f.pesado]);
      ws.getCell(`D${r}`).value = rec[f.cm1] || '';
      ws.getCell(`E${r}`).value = rec[f.cm2] || '';
      ws.getCell(`F${r}`).value = rec[f.cm3] || '';
      ws.getCell(`G${r}`).value = rec[f.resp] || '';
      ws.getCell(`H${r}`).value = (rec[f.fecha] instanceof Date) ? rec[f.fecha].toISOString().slice(0,10) : (rec[f.fecha] || '');
      ws.getCell(`I${r}`).value = rec[f.h1] || '';
      ws.getCell(`J${r}`).value = rec[f.h2] || '';
      ws.getCell(`K${r}`).value = rec[f.n] || '';
    });

    // Bordes del bloque principal B11..K16
    setBorder([11,2,16,11]);

    // Panel derecho: Equipos para Pesado (M..P) filas 12..18
    ws.mergeCells('M12:P12'); ws.getCell('M12').value = 'Equipos para Pesado:'; ws.getCell('M12').alignment = { horizontal:'center' }; ws.getCell('M12').font = { bold:true };
    const eqRows = [
      ['Balanza  74-M','equipo_balanza_74m'],
      ['Cámara flujo laminar 8-M','equipo_camara_8m'],
      ['Balanza  6-M','equipo_balanza_6m'],
      ['Mesón de  traspaso','equipo_meson_traspaso'],
      ['Balanza  99-M','equipo_balanza_99m'],
      ['Balanza  108-M','equipo_balanza_108m', true],
    ];
    eqRows.forEach((row,i) => {
      const rr = 13 + i;
      ws.mergeCells(`M${rr}:O${rr}`);
      ws.getCell(`M${rr}`).value = row[0];
      if (row[2]) ws.getCell(`M${rr}`).font = { bold:true };
      ws.getCell(`M${rr}`).alignment = { horizontal:'center' };
      ws.getCell(`P${rr}`).value = check(rec[row[1]]);
      ws.getCell(`P${rr}`).alignment = { horizontal:'center' };
    });
    setBorder([12,13,18,16]); // M12..P18

    // Panel derecho: Lugar de almacenamiento (R..T) filas 12..16
    ws.mergeCells('R12:T12'); ws.getCell('R12').value = 'Lugar de almacenamiento:'; ws.getCell('R12').alignment = { horizontal:'center' }; ws.getCell('R12').font = { bold:true };
    const locRows = [
      ['Frezeer 33-M','storage_freezer_33m'],
      ['Refrigerador 33-M','storage_refrigerador_33m'],
      ['Gabinete de Traspaso','storage_gabinete_traspaso'],
    ];
    locRows.forEach((row,i) => {
      const rr = 13 + i;
      ws.mergeCells(`R${rr}:S${rr}`);
      ws.getCell(`R${rr}`).value = row[0];
      ws.getCell(`R${rr}`).alignment = { horizontal:'center' };
      ws.getCell(`T${rr}`).value = check(rec[row[1]]);
      ws.getCell(`T${rr}`).alignment = { horizontal:'center' };
    });
    setBorder([12,18,16,20]); // R12..T16

  // ==============================
  // Sección: Puntas/pipetas desechables para pesado (*)
  // ==============================
  // Ubicamos este bloque como panel derecho bajo "Lugar de almacenamiento",
  // usando columnas R..T, para que coincida con el panel angosto de la imagen.
  const baseP = 18; // inicia dos filas debajo del panel derecho

  // Encabezado (dos filas) centrado y con fondo gris
  ws.mergeCells(`R${baseP}:T${baseP}`);
  ws.getCell(`R${baseP}`).value = 'Puntas/pipetas desechables para pesado';
  ws.getCell(`R${baseP}`).alignment = { horizontal:'center', vertical:'middle' };
  ws.getCell(`R${baseP}`).font = { bold:true };
  ws.getCell(`R${baseP}`).fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFD9D9D9'} };

  ws.mergeCells(`R${baseP+1}:T${baseP+1}`);
  ws.getCell(`R${baseP+1}`).value = '(*)';
  ws.getCell(`R${baseP+1}`).alignment = { horizontal:'center', vertical:'middle' };
  ws.getCell(`R${baseP+1}`).font = { bold:true };
  ws.getCell(`R${baseP+1}`).fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFD9D9D9'} };

  // Filas de datos: etiqueta en R, valores en S..T
  ws.getCell(`R${baseP+2}`).value = '1ml /clave:';
  ws.mergeCells(`S${baseP+2}:T${baseP+2}`);
  ws.getCell(`S${baseP+2}`).value = rec.clave_1ml || '';

  ws.getCell(`R${baseP+3}`).value = '10 ml/clave:';
  ws.mergeCells(`S${baseP+3}:T${baseP+3}`);
  ws.getCell(`S${baseP+3}`).value = rec.clave_10ml || '';

  // Área "Otros" con varias filas a la derecha
  ws.mergeCells(`R${baseP+4}:R${baseP+6}`);
  ws.getCell(`R${baseP+4}`).value = 'Otros:';
  ws.getCell(`R${baseP+4}`).alignment = { vertical:'middle' };
  ws.mergeCells(`S${baseP+4}:T${baseP+6}`);
  ws.getCell(`S${baseP+4}`).value = rec.clave_otros || '';
  ws.getCell(`S${baseP+4}`).alignment = { vertical:'top', wrapText:true };

  // Bordes alrededor del bloque completo R18..T24
  setBorder([baseP,18,baseP+6,20]);

    // ==============================
    // Pestañas adicionales: ram, RM y L, CT, CF y E.coli, sal, Entero, saureus
    // ==============================
    const extraTabs = ['ram', 'RM y L', 'CT, CF y E.coli', 'sal', 'Entero', 'saureus'];
    const headerMap = {
      'ram': {
        title: 'TRAZABILIDAD ANÁLISIS: ENUMERACIÓN DE AEROBIOS MESÓFILOS (NCh 2659.Of 2002)',
        code: 'R-PR-SVVM-M-4-11 / 15-02-23',
      },
      'default': {
        title: 'TRAZABILIDAD Y ANÁLISIS',
        code: 'R-INS-MM-M-1-15 /23-08-23',
      }
    };
    const renderHeaderOnly = (sheet, name) => {
      const conf = headerMap[name] || headerMap.default;
      // Título
      sheet.mergeCells('B1:T1');
      sheet.getCell('B1').value = conf.title;
      sheet.getCell('B1').alignment = { horizontal: 'center' };
      sheet.getCell('B1').font = { bold: true, size: 18 };
      // Código
      sheet.mergeCells('B2:T2');
      sheet.getCell('B2').value = conf.code;
      sheet.getCell('B2').alignment = { horizontal: 'center' };
      sheet.getCell('B2').font = { bold: true, size: 12 };
      // Banda ID
      sheet.mergeCells('I3:K3');
      const idCell2 = sheet.getCell('I3');
      idCell2.value = `${sample_id}`;
      idCell2.alignment = { horizontal: 'center', vertical: 'middle' };
      idCell2.font = { bold: true, size: 14 };
      idCell2.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFF2F2F2'} };
      sheet.getRow(3).height = 24;
      // Borde ID
      for (let c=9; c<=11; c++) {
        const cell = sheet.getCell(3,c);
        cell.border = { top:{style:'thin'},left:{style:'thin'},bottom:{style:'thin'},right:{style:'thin'} };
      }
      // Etiqueta ALI en H3
      const ali2 = sheet.getCell('H3');
      ali2.value = 'ALI';
      ali2.alignment = { horizontal:'center', vertical:'middle' };
      ali2.font = { bold:true };
      ali2.border = { top:{style:'thin'},left:{style:'thin'},bottom:{style:'thin'},right:{style:'thin'} };
    };
    extraTabs.forEach(name => {
      const s = wb.addWorksheet(name);
      renderHeaderOnly(s, name);
      s.getCell('B5').value = 'Sección en desarrollo';
      s.getCell('B5').font = { italic: true, color: { argb: 'FF808080' } };
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="TPA_${sample_id}.xlsx"`);
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    next(err);
  }
};
