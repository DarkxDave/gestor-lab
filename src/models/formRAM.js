const { query } = require('../db');
const samples = require('./samples');

// Columns present in form_ram_entries corresponding to form field names
const columns = [
  // Fechas y análisis
  'inicio_incubacion_fecha',
  'inicio_incubacion_hora',
  'inicio_incubacion_analista',
  'termino_analisis_fecha',
  'termino_analisis_hora',
  'termino_analisis_analista',
  // Control Ambiental
  'ca_pesado_temp',
  'ca_pesado_ufc',
  'ca_siembra',
  'ca_ecoli_ufc',
  'ca_blanco_ufc',
  // Siembra
  'siembra_tiempo_ok',
  'siembra_tiempo_minutos',
  'siembra_n_muestra_10g_90ml',
  'siembra_n_muestra_50g_450ml',
  // Controles de Calidad
  'cc_duplicado_ali_detalle',
  'cc_duplicado_ali_analisis',
  'cc_duplicado_ali_cumple',
  'cc_control_pos_blanco_ali_detalle',
  'cc_control_pos_blanco_ali_analisis',
  'cc_control_pos_blanco_ali_cumple',
  'cc_control_siembra_ali_detalle',
  'cc_control_siembra_ali_analisis',
  'cc_control_siembra_ali_cumple',
  // Control de Calidad 2
  'cc2_pesado_temp',
  'cc2_pesado_ufc',
  'cc2_siembra',
  'cc2_hora_inicio',
  'cc2_hora_termino',
  'cc2_temp',
  'cc2_ecoli_ufc',
  'cc2_blanco_ufc',
  // MIC
  'mic_desfavorable_si',
  'mic_desfavorable_no',
  'mic_tabla_pagina',
  'mic_limite',
  'mic_fecha_entrega',
  'mic_hora_entrega',
  // Muestrario
  'muestrario_muestra_rep_1',
  'muestrario_muestra_rep_2',
  'muestrario_dil_1',
  'muestrario_dil_2',
  'muestrario_c1_1',
  'muestrario_c1_2',
  'muestrario_c2_1',
  'muestrario_c2_2',
  'muestrario_sumc_1',
  'muestrario_sumc_2',
  'muestrario_d_1',
  'muestrario_d_2',
  'muestrario_n1_1',
  'muestrario_n1_2',
  'muestrario_n2_1',
  'muestrario_n2_2',
  'muestrario_x_1',
  'muestrario_x_2',
  'muestrario_resultado_ram_1',
  'muestrario_resultado_ram_2',
  'muestrario_resultado_rpes_1',
  'muestrario_resultado_rpes_2',
  // Notas
  'notes',
  'observaciones',
];

const booleanFields = new Set([
  'siembra_tiempo_ok',
  'mic_desfavorable_si',
  'mic_desfavorable_no',
]);

const tinyintFields = new Set([
  'cc_duplicado_ali_cumple',
  'cc_control_pos_blanco_ali_cumple',
  'cc_control_siembra_ali_cumple',
]);

function normalizeValue(key, val) {
  if (booleanFields.has(key)) {
    // Checkbox present => 'on' or any truthy => 1; else 0
    return val ? 1 : 0;
  }
  if (tinyintFields.has(key)) {
    if (val === undefined || val === null || val === '') return null;
    return String(val) === '1' ? 1 : 0;
  }
  if (key === 'siembra_tiempo_minutos') {
    if (val === undefined || val === null || val === '') return null;
    const n = Number(val);
    return Number.isFinite(n) ? n : null;
  }
  // Empty string to null to avoid storing empty strings
  if (val === undefined || val === null || val === '') return null;
  return val;
}

exports.save = async (sample_id, data = {}) => {
  await samples.ensureSample(sample_id);

  const vals = columns.map((c) => normalizeValue(c, data[c]));

  const colList = columns.join(', ');
  const placeholders = columns.map(() => '?').join(', ');
  const updates = columns.map((c) => `${c}=VALUES(${c})`).join(', ');

  const sql = `
    INSERT INTO form_ram_entries (sample_id, ${colList}, created_at, updated_at)
    VALUES (?, ${placeholders}, NOW(), NOW())
    ON DUPLICATE KEY UPDATE ${updates}, updated_at=NOW()
  `;

  await query(sql, [sample_id, ...vals]);
};

exports.getBySampleId = async (sample_id) => {
  const rows = await query('SELECT * FROM form_ram_entries WHERE sample_id=?', [sample_id]);
  return rows[0] || null;
};

// List for export summary: join samples with RAM
exports.listAll = async () => {
  return await query(`
    SELECT s.sample_id,
           ram.*
    FROM samples s
    LEFT JOIN form_ram_entries ram ON ram.sample_id = s.sample_id
    ORDER BY s.id DESC
  `);
};
