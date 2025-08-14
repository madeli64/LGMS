# Test Plan — LabelGen

## Acceptance tests
- Placeholders filled; `TXT_ITEM_NO` padded (e.g., 0001/0060)
- QR = `SKU|SERIAL` scanned OK (3 samples)
- PDF page size matches template; fonts embedded; ≥300 DPI
- Outputs under `Output/<job_id>/`; job JSON moved to `done/`

## Scenarios
1) Single label (count=1) all fields
2) 60× labels — padding & unique QR
3) Concurrency — 5 jobs × 20 labels
4) Errors — missing template / wrong shape name / QR failure → `failed/`
5) Font rendering on second machine

## KPI/SLA targets
- ≤20s for 10 labels; ≤90s for 100 labels
- ≥99% success in 1-week pilot
