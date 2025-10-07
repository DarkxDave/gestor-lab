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
