-- Special Transportation Listing: dataset + table schema (BigQuery)
-- Data is loaded already de-identified; see ../README.md.

CREATE SCHEMA IF NOT EXISTS `school-district-work.special_transportation`
OPTIONS (location = 'US');

CREATE TABLE IF NOT EXISTS `school-district-work.special_transportation.special_transportation_listing_synthetic`
(
  snapshot_date                  DATE   OPTIONS(description = 'Date of the source listing export'),
  school                         STRING OPTIONS(description = 'Generic school label; real school names removed'),
  school_level                   STRING OPTIONS(description = 'Elementary / Middle / High / Other'),
  pp_school_agency               STRING OPTIONS(description = 'Generic parentally-placed school/agency label'),
  last_name                      STRING OPTIONS(description = 'Synthetic last name'),
  first_name                     STRING OPTIONS(description = 'Synthetic first name'),
  student_id                     INT64  OPTIONS(description = 'Synthetic 9-digit ID, stable across snapshots; NULL where the source ID# was blank'),
  need                           STRING,
  transportation_recommendation  STRING,
  cr_decision_status             STRING,
  di_finalized                   STRING OPTIONS(description = 'Yes/No; NULL for the June snapshots'),
  di_date_finalized              DATE   OPTIONS(description = 'NULL for the June snapshots and where DI Finalized = No')
)
CLUSTER BY snapshot_date, school;
