-- Deprecated: This project installs clean from scripts/init_db.sql.
-- This migration has been superseded by the baseline and is intentionally left empty.

-- Migration: Remove Muestrario columns from form_ram_entries
ALTER TABLE form_ram_entries
  DROP COLUMN IF EXISTS muestrario_muestra_rep_1,
  DROP COLUMN IF EXISTS muestrario_muestra_rep_2,
  DROP COLUMN IF EXISTS muestrario_dil_1,
  DROP COLUMN IF EXISTS muestrario_dil_2,
  DROP COLUMN IF EXISTS muestrario_c1_1,
  DROP COLUMN IF EXISTS muestrario_c1_2,
  DROP COLUMN IF EXISTS muestrario_c2_1,
  DROP COLUMN IF EXISTS muestrario_c2_2,
  DROP COLUMN IF EXISTS muestrario_sumc_1,
  DROP COLUMN IF EXISTS muestrario_sumc_2,
  DROP COLUMN IF EXISTS muestrario_d_1,
  DROP COLUMN IF EXISTS muestrario_d_2,
  DROP COLUMN IF EXISTS muestrario_n1_1,
  DROP COLUMN IF EXISTS muestrario_n1_2,
  DROP COLUMN IF EXISTS muestrario_n2_1,
  DROP COLUMN IF EXISTS muestrario_n2_2,
  DROP COLUMN IF EXISTS muestrario_x_1,
  DROP COLUMN IF EXISTS muestrario_x_2,
  DROP COLUMN IF EXISTS muestrario_resultado_ram_1,
  DROP COLUMN IF EXISTS muestrario_resultado_ram_2,
  DROP COLUMN IF EXISTS muestrario_resultado_rpes_1,
  DROP COLUMN IF EXISTS muestrario_resultado_rpes_2;
