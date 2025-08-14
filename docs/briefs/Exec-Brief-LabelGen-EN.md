# Executive Brief — LabelGen (PowerPoint→PDF) with WMS Readiness
Version: 1.0.0
Audience: Maya (Ops) & Vincent (OOP)
Owner: Mohsen Adelimoghaddam
Status: Approval Requested (Pilot)

## 1) Summary
We propose a lightweight system that converts designer-owned PowerPoint templates into print-ready PDF labels. Variable fields are filled from a simple form; each label includes a **unique QR** (payload `SKU|SERIAL`) and a padded item number like `0001/0060`.  
**Why now:** Our previous **Form → PAC → SVG/HTML → PDF** prototype showed layout drift and slow rendering. Moving to native PowerPoint rendering eliminates those issues while staying within Microsoft 365.

## 2) Why switch (from SVG/HTML/PDF)
| Issue in old path | Impact | New path (PowerPoint→PDF) |
|---|---|---|
| CSS/HTML layout shifts across renderers | Misaligned labels | PowerPoint engine preserves exact designer layout |
| Font fallback/embedding quirks | Brand inconsistency | Office embeds licensed fonts reliably |
| Server-side rasterization latency | Slow batch jobs | Local agent with COM is fast (≤20s for 10 labels; ≤90s for 100) |
| Complex HTML templating | Ongoing maintenance | Designers work directly in PPT; devs just fill placeholders |
| Image DPI & scaling | Blurry prints | Exact slide size = label size; ≥300 DPI export |

## 3) What we will deliver (MVP scope)
- **E2E pipeline:** Power Apps form → PAC writes Job JSON → PAD agent fills PPT (text/images/QR) → **PDF in SharePoint**.
- **Designer freedom:** Mark text as `{{product_name}}`, images via picture **Alt Text** `{{img:logo}}`/`{{img:qr}}`; our Profiler auto-standardizes to `TXT_*` / `IMG_*`.
- **Per-job outputs:** PDF, status JSON, thumbnail under `ServerPAD/Output/<job_id>/`.
- **Duo language ready (EN/FR)** — no Persian fonts required.
- **Out of scope (explicit):** **Printing** and external routing.

## 4) Access & environment (asks)
1) **R/W access to the `ServerPAD` library** for the PAD agent account.  
2) One small **PC/VM** (always-on) with **Power Automate Desktop** + **PowerPoint** installed.  
3) Confirm **printing is out of scope** in this phase.  
4) Nominate a **Template Owner** (designer) for layout changes.

## 5) Timeline (pilot)
- **Week 0–1:** Environment & two sample templates (4×6 in, 100×150 mm).  
- **Week 2:** E2E pilot on 2 SKUs (scenarios: 1×, 20×, 60×).  
- **Week 3:** KPI report + go/no-go and (optionally) WMS alignment kickoff.

## 6) Success criteria (KPI/Acceptance)
- All placeholders filled; `TXT_ITEM_NO` padded (`0001/0060`).  
- QR = `SKU|SERIAL`, scans OK on samples.  
- SLA: ≤**20s** (10 labels), ≤**90s** (100 labels).  
- ≥**99%** success rate during pilot week.  
- PDFs stored under `ServerPAD/Output/<job_id>/`; jobs moved to `done/`.

## 7) Risks & mitigations
- **PowerPoint COM reliability:** run on a dedicated PC/VM; watchdog & restart; logs in `ServerPAD/Logs`.  
- **Template naming mistakes:** Profiler generates a profile file; form shows required fields; validation before run.  
- **SharePoint sync hiccups:** atomic move `incoming→processing→done/failed`; clear runbook to retry.  
- **Future central serials:** we start local (`{JOB8}-{SEQ}`) and can switch to central allocation without template changes.  
- **Security/permissions:** minimum R/W to the `ServerPAD` library; no PII in labels.

## 8) Cost & licensing
- **No new licenses.** Uses Microsoft 365 + existing PowerPoint/PAD.  
- Hardware: one small PC/VM (existing or low-cost).

## 9) WMS alignment (next step, not in MVP)
- Keep QR payload `SKU|SERIAL`. Later, a central **range allocator** (Flow + SharePoint List) can issue organization-wide serials (e.g., `PLT-{SKU6}-{YYWW}-{NUM}`) without changing templates or QR parsing.  
- Minimal WMS tables ready (Serials, Ranges, LabelJobs/Labels).

## Decision needed
- ✅ Approve the **pilot** with the access & environment above.  
- ✅ Confirm **printing** is out of scope for MVP.  
- ✅ Assign a **Template Owner**.

