-- Crear base de datos y tablas para Gestor Lab
CREATE DATABASE IF NOT EXISTS gestor_lab CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci;
USE gestor_lab;

-- Tabla de muestras
CREATE TABLE IF NOT EXISTS samples (
  id INT AUTO_INCREMENT PRIMARY KEY,
  sample_id VARCHAR(100) NOT NULL UNIQUE
) ENGINE=InnoDB;

-- Formulario A (Trazabilidad - TPA)
DROP TABLE IF EXISTS form_a_entries;
CREATE TABLE form_a_entries (
  id INT AUTO_INCREMENT PRIMARY KEY,
  sample_id VARCHAR(100) NOT NULL,
  -- Almacenamiento
  storage_freezer_33m BOOLEAN NOT NULL DEFAULT 0,
  storage_refrigerador_33m BOOLEAN NOT NULL DEFAULT 0,
  storage_meson_siembra BOOLEAN NOT NULL DEFAULT 0,
  storage_gabinete_traspaso BOOLEAN NOT NULL DEFAULT 0,
  observaciones TEXT NULL,
  -- Manipulación de muestras (1..3)
  retiro_muestra_1 BOOLEAN NOT NULL DEFAULT 0,
  pesado_1 BOOLEAN NOT NULL DEFAULT 0,
  clave_material_1 VARCHAR(60) NULL,
  clave_material_2 VARCHAR(60) NULL,
  clave_material_3 VARCHAR(60) NULL,
  responsable_1 VARCHAR(255) NULL,
  fecha_1 DATE NULL,
  hora_inicio_1 TIME NULL,
  hora_termino_1 TIME NULL,
  n_muestra_1 VARCHAR(50) NULL,

  retiro_muestra_2 BOOLEAN NOT NULL DEFAULT 0,
  pesado_2 BOOLEAN NOT NULL DEFAULT 0,
  clave_material_4 VARCHAR(60) NULL,
  clave_material_5 VARCHAR(60) NULL,
  clave_material_6 VARCHAR(60) NULL,
  responsable_2 VARCHAR(255) NULL,
  fecha_2 DATE NULL,
  hora_inicio_2 TIME NULL,
  hora_termino_2 TIME NULL,
  n_muestra_2 VARCHAR(50) NULL,

  retiro_muestra_3 BOOLEAN NOT NULL DEFAULT 0,
  pesado_3 BOOLEAN NOT NULL DEFAULT 0,
  clave_material_7 VARCHAR(60) NULL,
  clave_material_8 VARCHAR(60) NULL,
  clave_material_9 VARCHAR(60) NULL,
  responsable_3 VARCHAR(255) NULL,
  fecha_3 DATE NULL,
  hora_inicio_3 TIME NULL,
  hora_termino_3 TIME NULL,
  n_muestra_3 VARCHAR(50) NULL,

  -- Equipos para pesado y lugar de almacenamiento
  equipo_balanza_74m BOOLEAN NOT NULL DEFAULT 0,
  equipo_camara_8m BOOLEAN NOT NULL DEFAULT 0,
  equipo_balanza_6m BOOLEAN NOT NULL DEFAULT 0,
  equipo_meson_traspaso BOOLEAN NOT NULL DEFAULT 0,
  equipo_balanza_99m BOOLEAN NOT NULL DEFAULT 0,
  equipo_balanza_108m BOOLEAN NOT NULL DEFAULT 0,
  equipo_freezer_33m BOOLEAN NOT NULL DEFAULT 0,
  equipo_refrigerador_33m BOOLEAN NOT NULL DEFAULT 0,
  equipo_gabinete_traspaso BOOLEAN NOT NULL DEFAULT 0,

  -- Micropipetas y limpieza
  micropipeta_22m BOOLEAN NOT NULL DEFAULT 0,
  micropipeta_32m BOOLEAN NOT NULL DEFAULT 0,
  micropipeta_23m BOOLEAN NOT NULL DEFAULT 0,
  micropipeta_75m BOOLEAN NOT NULL DEFAULT 0,
  micropipeta_72m BOOLEAN NOT NULL DEFAULT 0,
  micropipeta_94m BOOLEAN NOT NULL DEFAULT 0,
  micropipeta_98m BOOLEAN NOT NULL DEFAULT 0,
  micropipeta_103m BOOLEAN NOT NULL DEFAULT 0,
  micropipeta_100m BOOLEAN NOT NULL DEFAULT 0,
  micropipeta_102m BOOLEAN NOT NULL DEFAULT 0,
  micropipeta_106m BOOLEAN NOT NULL DEFAULT 0,
  clave_1ml VARCHAR(100) NULL,
  clave_10ml VARCHAR(100) NULL,
  clave_otros VARCHAR(255) NULL,

  limpieza_meson BOOLEAN NOT NULL DEFAULT 0,
  limpieza_stomacher BOOLEAN NOT NULL DEFAULT 0,
  limpieza_camara BOOLEAN NOT NULL DEFAULT 0,
  limpieza_balanza BOOLEAN NOT NULL DEFAULT 0,
  limpieza_balanza2 BOOLEAN NOT NULL DEFAULT 0,
  limpieza_otros VARCHAR(255) NULL,
  limpieza_aerosol BOOLEAN NOT NULL DEFAULT 0,
  observaciones_limpieza TEXT NULL,

  -- Siembra y diluyentes
  clave_general VARCHAR(255) NULL,
  clave_puntas_1ml VARCHAR(100) NULL,
  bano_5m BOOLEAN NOT NULL DEFAULT 0,
  clave_puntas_10ml VARCHAR(100) NULL,
  homogenizador_12m BOOLEAN NOT NULL DEFAULT 0,
  clave_placas VARCHAR(100) NULL,
  cuenta_colonias_9m BOOLEAN NOT NULL DEFAULT 0,
  clave_asas VARCHAR(100) NULL,
  cuenta_colonias_101m BOOLEAN NOT NULL DEFAULT 0,
  clave_blender VARCHAR(100) NULL,
  phmetro_93m BOOLEAN NOT NULL DEFAULT 0,
  clave_bolsas VARCHAR(100) NULL,
  pipetas_desechables BOOLEAN NOT NULL DEFAULT 0,
  clave_probeta VARCHAR(100) NULL,
  clave_otro VARCHAR(255) NULL,

  ap_90ml VARCHAR(100) NULL,
  ap_tubos_ml VARCHAR(100) NULL,
  ap_99ml VARCHAR(100) NULL,
  sps_225ml VARCHAR(100) NULL,
  ap_450ml VARCHAR(100) NULL,
  sps_tubos VARCHAR(100) NULL,
  ap_225ml VARCHAR(100) NULL,
  sps_sa_90ml VARCHAR(100) NULL,
  ap_500ml VARCHAR(100) NULL,
  sps_sa_tubos VARCHAR(100) NULL,
  pbs_450ml VARCHAR(100) NULL,
  diluyente_otro VARCHAR(255) NULL,
  diluyente_otros1 VARCHAR(255) NULL,

  created_at DATETIME NOT NULL,
  updated_at DATETIME NOT NULL,
  CONSTRAINT fk_form_a_sample FOREIGN KEY (sample_id) REFERENCES samples(sample_id) ON DELETE CASCADE ON UPDATE CASCADE
) ENGINE=InnoDB;

-- Formulario B
CREATE TABLE IF NOT EXISTS form_b_entries (
  id INT AUTO_INCREMENT PRIMARY KEY,
  sample_id VARCHAR(100) NOT NULL,
  notes TEXT NULL,
  approved BOOLEAN NOT NULL DEFAULT 0,
  qc_pass BOOLEAN NOT NULL DEFAULT 0,
  created_at DATETIME NOT NULL,
  updated_at DATETIME NOT NULL,
  CONSTRAINT fk_form_b_sample FOREIGN KEY (sample_id) REFERENCES samples(sample_id) ON DELETE CASCADE ON UPDATE CASCADE
) ENGINE=InnoDB;
