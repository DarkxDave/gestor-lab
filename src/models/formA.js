const { query } = require('../db');
const samples = require('./samples');

const booleanFields = [
  'storage_freezer_33m','storage_refrigerador_33m','storage_meson_siembra','storage_gabinete_traspaso',
  'retiro_muestra_1','pesado_1','retiro_muestra_2','pesado_2','retiro_muestra_3','pesado_3',
  'equipo_balanza_74m','equipo_camara_8m','equipo_balanza_6m','equipo_meson_traspaso','equipo_balanza_99m','equipo_balanza_108m','equipo_freezer_33m','equipo_refrigerador_33m','equipo_gabinete_traspaso',
  'micropipeta_22m','micropipeta_32m','micropipeta_23m','micropipeta_75m','micropipeta_72m','micropipeta_94m','micropipeta_98m','micropipeta_103m','micropipeta_100m','micropipeta_102m','micropipeta_106m',
  'limpieza_meson','limpieza_stomacher','limpieza_camara','limpieza_balanza','limpieza_balanza2','limpieza_aerosol',
  'bano_5m','homogenizador_12m','cuenta_colonias_9m','cuenta_colonias_101m','phmetro_93m','pipetas_desechables'
];

const textFields = [
  'observaciones',
  'clave_material_1','clave_material_2','clave_material_3','responsable_1','n_muestra_1',
  'clave_material_4','clave_material_5','clave_material_6','responsable_2','n_muestra_2',
  'clave_material_7','clave_material_8','clave_material_9','responsable_3','n_muestra_3',
  'clave_1ml','clave_10ml','clave_otros',
  'limpieza_otros','observaciones_limpieza',
  'clave_general','clave_puntas_1ml','clave_puntas_10ml','clave_placas','clave_asas','clave_blender','clave_bolsas','clave_probeta','clave_otro',
  'ap_90ml','ap_tubos_ml','ap_99ml','sps_225ml','ap_450ml','sps_tubos','ap_225ml','sps_sa_90ml','ap_500ml','sps_sa_tubos','pbs_450ml','diluyente_otro','diluyente_otros1'
];

const dateFields = ['fecha_1','fecha_2','fecha_3'];
const timeFields = ['hora_inicio_1','hora_termino_1','hora_inicio_2','hora_termino_2','hora_inicio_3','hora_termino_3'];

exports.save = async (sample_id, payload) => {
  await samples.ensureSample(sample_id);
  const rows = await query('SELECT id FROM form_a_entries WHERE sample_id = ?', [sample_id]);

  // Normalize booleans and nullables
  const data = {};
  booleanFields.forEach(k => data[k] = !!payload[k]);
  textFields.forEach(k => data[k] = payload[k] ? String(payload[k]) : null);
  dateFields.forEach(k => data[k] = payload[k] || null);
  timeFields.forEach(k => data[k] = payload[k] || null);

  const columns = [
    ...booleanFields,
    ...textFields,
    ...dateFields,
    ...timeFields,
  ];
  const values = columns.map(k => data[k]);

  if (rows.length) {
    const setClause = columns.map(c => `${c}=?`).join(', ');
    await query(`UPDATE form_a_entries SET ${setClause}, updated_at=NOW() WHERE sample_id=?`, [...values, sample_id]);
  } else {
    const colList = columns.join(', ');
    const placeholders = columns.map(() => '?').join(', ');
    await query(`INSERT INTO form_a_entries (sample_id, ${colList}, created_at, updated_at) VALUES (?, ${placeholders}, NOW(), NOW())`, [sample_id, ...values]);
  }
};

exports.getBySampleId = async (sample_id) => {
  const rows = await query('SELECT * FROM form_a_entries WHERE sample_id = ?', [sample_id]);
  if (!rows[0]) return null;
  return rows[0];
};

exports.listJoinFormB = async () => {
  return await query(`
    SELECT s.sample_id,
           a.*,
           b.notes AS b_notes, b.approved AS b_approved, b.qc_pass AS b_qc_pass
    FROM samples s
    LEFT JOIN form_a_entries a ON a.sample_id = s.sample_id
    LEFT JOIN form_b_entries b ON b.sample_id = s.sample_id
    ORDER BY s.id DESC
  `);
};
