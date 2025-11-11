-- Deprecated: This project installs clean from scripts/init_db.sql.
-- This migration has been superseded by the baseline and is intentionally left empty.

-- Migration: add new Datos fields to form_ram_entries
-- Note: If columns already exist, this migration may fail; run conditionally if using MySQL 8.0.29+ with IF NOT EXISTS.

ALTER TABLE form_ram_entries
  ADD COLUMN datos_suspension_inicial_den VARCHAR(50) NULL AFTER mic_hora_entrega,
  ADD COLUMN datos_volumen_petri_ml VARCHAR(50) NULL AFTER datos_suspension_inicial_den;
