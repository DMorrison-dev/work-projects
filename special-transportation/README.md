# Special Transportation Listing — De-identified BigQuery Project

Snapshot analysis of a Special Education special-transportation listing: which transportation
needs and recommendations are on file, how they change between exports, and how complete the
finalization (DI) step is.

**All data is de-identified. No real student information is stored in BigQuery or in this repo.**

| | |
|---|---|
| BigQuery project | `school-district-work` |
| Dataset | `special_transportation` (kept separate from the other datasets) |
| Table | `special_transportation_listing_synthetic` |
| Grain | one row per student + need + transportation recommendation, per snapshot |
| Rows | 11,236 across 4 snapshots |

## Snapshots

| snapshot_date | Rows | Distinct student IDs |
|---|---|---|
| 06/01/2026 | 2,698 | 1,851 |
| 06/08/2026 | 2,699 | 1,852 |
| 08/25/2026 | 2,919 | 2,034 |
| 09/08/2026 | 2,920 | 2,027 |

## De-identification

| Source field | Treatment |
|---|---|
| Last Name / First Name | Replaced with random generic names |
| ID# | Replaced with a random synthetic 9-digit ID (`9xxxxxxxx`) |
| School | Replaced with generic labels (`Elementary School 04`, `Program Site 07`, ...) |
| PP School Agency | Replaced with generic labels (`Outside Agency 06`); `Home Schooled` kept |
| Need, Transportation Recommendation, CR Decision/Status, DI Finalized, DI Date Finalized | Unchanged |

- A student keeps the same synthetic name and ID in every snapshot, so changes over time can still be tracked.
- `school_level` (Elementary / Middle / High / Other) is derived from keywords in the source school name.
- The real-to-synthetic mapping was deleted after the load; the table cannot be matched back to real students.
- Source files and data extracts are never committed (`*.xlsx` / `*.csv` are git-ignored).

## Known data notes

- The two June snapshots do not carry the DI columns, so `di_finalized` / `di_date_finalized` are NULL there.
- Five rows (two on 08/25, three on 09/08) have a blank school and ID# in the source; they are kept blank.

## Files

```
sql/01_create_table.sql          table schema (DDL)
sql/02_data_quality_checks.sql   row counts, duplicates, nulls, DI consistency, ID format
sql/03_analysis_queries.sql      needs mix, school-level summary, finalization rate, snapshot-over-snapshot changes
```
