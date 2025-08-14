# SPEC-1-LabelGen — PDF Label Generation from PowerPoint in Microsoft Business

Version: 1.0.0 — 2025-08-14  
Status: Approved (MVP locked: A/local; Future B/central)  
Owner: <Owner/Team>

<!-- Paste the full EN spec content here (from your Master canvas). -->
# SPEC-1-LabelGen — PDF Label Generation from PowerPoint in Microsoft Business

**Version:** 1.0.0 — 2025-08-14
**Master:** YES (English)
**Status:** MVP locked (**Serial A/Local**), Future **B/Central** planned
**Decisions summary:** Multi-template (PPTX) • JSON-driven dynamic **fields/images** • **Profiler** discovers `{{field}}` & `{{img:name}}` from **Text + Alt Text** • QR payload = `SKU|SERIAL` • `TXT_ITEM_NO` with **dynamic padding** (e.g., `0001/0060`) • Per-job outputs (PDF/status/thumb) • PAC → `Jobs/incoming` • PAD moves atomically `incoming→processing→done/failed` • Retention **90/30/14** • SLA **20s(10)** / **90s(100)**

---

## Background

* **Problem:** Generate tamper-proof **PDF labels** for boxes from designer-owned **PowerPoint templates**. Variable parts are auto-filled by naming shapes (Selection Pane) and substituting values.
* **Target flow:**

  1. User submits a form in **Power Apps** → **Power Automate Cloud (PAC)** writes a **Job JSON** into **SharePoint** Document Library `ServerPAD` (`Jobs/incoming/`).
  2. A Windows agent running **Power Automate Desktop (PAD)** polls every 5s, moves new jobs to `processing/` **atomically**, and runs a local script.
  3. **LabelGen.ps1** (PowerShell + PowerPoint COM + QRCoder.dll) fills text/images, generates **unique QR per slide**, exports **PDF**, writes outputs under `Output/<job_id>/`, returns status.
* **Constraints:** EN/FR environment; always-on Windows agent with PowerPoint; idempotency & minimal concurrency; versioned templates; print-grade fonts embedded.

---

## Requirements (MoSCoW)

**Must**

* Accept Job JSON with: `job_id`, `template_path`, `language`, `sku`, **either** `labels_count` or `labels[]`, `serialProvider`, `qr`.
* Create **N slides** per job (1 label/slide) with `TXT_ITEM_NO = i/N` (padded) and **unique QR** per slide: payload `SKU|SERIAL`.
* Map text to `TXT_*` shapes; images to `IMG_*` shapes. Load template from SharePoint sync, export **PDF** with embedded fonts, store in same library.
* Idempotency via `job_id`; move JSON to `processing/done/failed`; log execution & errors; return status to Power Apps/Flow.

**Should**

* Multi-template/version support and version validation.
* JSON Schema validation.
* Standard output naming.
* Concurrency for \~5 jobs/min; job up to 20 labels.
* QR settings configurable (ECC, pixelsPerModule, margin).

**Could**

* First-page **thumbnail** (PNG) for preview.
* Daily CSV stats.
* Future support for other barcodes (Code128/DataMatrix).

**Won’t (now)**

* Printing or external routing.
* Non-PowerPoint renderers.

---

## Method

### 1) Job JSON (dynamic fields + images)

```
{
  "job_id": "9b2c6a7e-6c2e-4c7e-9e1e-12c0a3f1a001",
  "template_path": "Templates/label_4x6in_v2.pptx",
  "language": "en-CA",
  "sku": "DLP-OM300-VU-GV-40-15D-M200-70-LN-BTU-DI",
  "labels_count": 60,
  "serialProvider": { "source": "local", "format": "{JOB}-{SEQ:000000}", "start": 1 },
  "qr": { "payload": "SKU|SERIAL", "ecc": "Q", "pixelsPerModule": 10, "margin": 2 },
  "fields": {
    "TXT_PRODUCT_NAME": "Widget X",
    "TXT_CLIENT_ADDRESS": "123 Rue Example\nMontreal, QC",
    "TXT_ITEM_NO": "{I}/{N}"
  },
  "images": {
    "IMG_LOGO": "Templates/assets/logo.png"
  }
}

```

* `fields` and `images` are **fully dynamic** (keys = shape names).
* System tokens in text values: `{I}`, `{N}`, `{SKU}`, `{SERIAL}` → replaced at generation.

### 2) Template Profiler + Image placeholders via Alt Text

**Designer contract**

* **Text placeholders (in text boxes):** `{{product_name}}`, `{{client_address}}` → become `TXT_PRODUCT_NAME`, `TXT_CLIENT_ADDRESS`.
* **Image placeholders:** insert a **sample image** (dummy logo or QR), then set **Alt Text** to `{{img:logo}}` or `{{img:qr}}` → become `IMG_LOGO`, `IMG_QR`.
* **System tokens (not double-braced):** `{I}`, `{N}`, `{SKU}`, `{SERIAL}`.
* Escape real braces with `\{\{` and `\}\}`.
* **Don’t group** pictures; keep “Compress Pictures” off for print quality.

**Profiler behavior**

* Scans both **shape text** and **Alt Text** for `{{...}}`.
* Renames shapes to `TXT_*` or `IMG_*`, ensuring uniqueness with `__2`, `__3`, …
* Writes a profile next to the template: `Templates/<name>.profile.json`, listing fields/images and their **shape lists** (so repeated fields fill everywhere).
* All fields are **optional by default**; your form UI can toggle `required` per profile entry.

**Example profile**

```
{
  "template": "Templates/label_4x6in_v2.pptx",
  "version": "v2",
  "fields": [
    { "name": "PRODUCT_NAME", "type": "text", "label": "Product name",
      "required": false, "multiline": false, "hideIfEmpty": false,
      "shapes": ["TXT_PRODUCT_NAME","TXT_PRODUCT_NAME__2"] },
    { "name": "CLIENT_ADDRESS", "type": "text", "label": "Client address",
      "required": false, "multiline": true, "hideIfEmpty": true,
      "shapes": ["TXT_CLIENT_ADDRESS"] }
  ],
  "images": [
    { "name": "LOGO", "type": "image", "required": false, "shapes": ["IMG_LOGO"] },
    { "name": "QR",   "type": "image", "required": false, "shapes": ["IMG_QR"] }
  ],
  "systemTokens": ["I","N","SKU","SERIAL"],
  "defaults": { "TXT_ITEM_NO": "{I}/{N}" }
}

```

### 3) Algorithm (pseudo)

```
Load job JSON → Open PPT template
If profile exists: validate required, apply defaults, build field/image maps

For i = 1..N:
  serial  = (A/local): "{JOB8}-{SEQ:000000}"  // JOB8 = first 8 hex of job_id
  payload = SKU + "|" + serial
  qrPng   = Generate QR (ECC=Q, ppm=10, margin=2)
  slide   = (i==1 ? first slide : duplicate)
  Fill text: replace {I}/{N}/{SKU} in values; set on all matching shapes
  Place IMG_QR from qrPng (Fill.UserPicture; fallback AddPicture preserving name)
  Place other images from job.images per profile
  Set TXT_ITEM_NO = padded i/N (e.g., 0001/0060)

Export PDF → Export thumbnail → Write status.json
Move job JSON to done/failed

```

### 4) PowerShell — key functions used by generator

```
function Set-Text($slide, $name, $text) {
  foreach ($s in @($slide.Shapes)) {
    if ($s.Name -eq $name -and $s.HasTextFrame) { $s.TextFrame.TextRange.Text = $text; return }
  }
}

# Keep placeholder name; if Fill fails, overlay a new picture and delete the sample
function Place-Image($slide, $name, $imgPath) {
  foreach ($s in @($slide.Shapes)) {
    if ($s.Name -eq $name) {
      try { $s.Fill.UserPicture($imgPath); return } catch {}
      $new = $slide.Shapes.AddPicture($imgPath, $false, $true, $s.Left, $s.Top, $s.Width, $s.Height)
      $new.Name = $s.Name
      $s.Delete()
      return
    }
  }
}

```

### 5) PlantUML — sequence

```
@startuml
actor User
participant "Power Apps" as PA
participant "Power Automate Cloud" as PAC
participant "SharePoint (ServerPAD)" as SP
participant "PAD Agent" as PAD
participant "LabelGen.ps1" as LG
participant "PowerPoint COM" as PPT
participant "QRCoder.dll" as QR

User -> PA: Submit form (sku/lang/count/template/fields)
PA -> PAC: Trigger
PAC -> SP: Create Job JSON (Jobs/incoming)
PAD -> SP: Poll & move to processing
PAD -> LG: Run LabelGen.ps1 -JobJsonPath
LG -> LG: Generate serials (JOB8-SEQ)
LG -> QR: Make QR PNG (per slide)
LG -> PPT: Open template / duplicate slide
LG -> PPT: Fill TXT_* and IMG_QR (+ images)
LG -> PPT: Export PDF + thumbnail
LG -> SP: Output/<job_id>/... + status.json
PAD -> SP: Move job to done/failed
PAC -> PA: Return status/url
@enduml

```

### 6) Serial strategy (and WMS alignment)

* **MVP (A/local):** `SERIAL = {JOB8}-{SEQ:000000}`; **QR** is always `SKU|SERIAL`.
* **Future (B/central):** allocate ranges via Flow + SharePoint List; recommended **format** `PLT-{SKU6}-{YYWW}-{NUM:00000}` (plant code, short SKU hash, year-week, sequence). No template/QR change needed.
* Target length ≤ 16 chars for readability on small labels; avoid ambiguous letters in human-readable strings.

---

## Implementation

### A) Folder structure (SharePoint + local sync)

```
ServerPAD/
 ├─ Cache/
 ├─ Jobs/
 │   ├─ incoming/      # PAC writes JSON here
 │   ├─ processing/    # PAD moves to <job_id>.lock.json here
 │   ├─ done/
 │   └─ failed/
 ├─ Logs/
 ├─ Output/
 │   └─ <job_id>/      # PDF + status.json + thumb.png
 ├─ Templates/         # PPTX + *.profile.json + assets/
 └─ Tools/
     └─ LabelGen/
         ├─ LabelGen.ps1
         └─ QRCoder.dll

```

> Local synced path example: `C:\Users\Learner\LightBase\Lightbase-Platform - ServerPAD`

### B) JSON Schema (Draft 2020-12; trimmed)

```
{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "type": "object",
  "required": ["job_id", "template_path", "language", "sku"],
  "properties": {
    "job_id": {"type":"string","format":"uuid"},
    "template_path": {"type":"string","pattern":"^Templates\\/.+\\.pptx$"},
    "language": {"type":"string","enum":["en-CA","fr-CA"]},
    "sku": {"type":"string","minLength":1},
    "labels_count": {"type":"integer","minimum":1},
    "labels": {"type":"array","items":{"type":"object","properties":{"serial":{"type":"string"}}}},
    "serialProvider": {
      "type":"object",
      "properties":{
        "source":{"type":"string","enum":["local","central"]},
        "format":{"type":"string"},
        "start":{"type":"integer","minimum":1}
      },
      "required":["source","format"]
    },
    "qr": {
      "type":"object",
      "properties":{
        "payload":{"const":"SKU|SERIAL"},
        "ecc":{"type":"string","enum":["L","M","Q","H"],"default":"Q"},
        "pixelsPerModule":{"type":"integer","minimum":3,"default":10},
        "margin":{"type":"integer","minimum":0,"default":2}
      },
      "required":["payload"]
    },
    "fields":{"type":"object"},
    "images":{"type":"object"}
  },
  "oneOf":[{"required":["labels_count"]},{"required":["labels"]}]
}

```

### C) PAD flow (variables & steps)

* **Variables**:
  `varIncoming, varProcessing, varDone, varFailed, varOutput, varTemplates, rootShare("%SERVERPAD%")`
* **Loop**

  1. Get files from `Jobs/incoming` → if none: Wait 5s.
  2. For each file: **Move** → `Jobs/processing/<job_id>.lock.json`.
  3. **Run PowerShell**:

     ```
     powershell.exe -ExecutionPolicy Bypass -File "%SERVERPAD%\Tools\LabelGen\LabelGen.ps1" -JobJsonPath "%jobMovedPath%" -Root "%SERVERPAD%"

     ```
  4. If exit code **0** → move JSON to **done/**; else → **failed/** and log.

### D) PAC flow (from Power Apps)

* **Trigger:** Power Apps (params: `product_variant_code`, `language`, `count`, `templateFilename`, `fieldsJson`)
* **Compose Job** (as in JSON example); create file at `ServerPAD/Jobs/incoming/<job_id>.json`; respond `{ status:'created', jobUrl:'...' }`.

### E) Power Apps (Power Fx snippets)

**Templates dropdown (from ****\`\`****):**

```
ddTemplate.Items =
Sort(
  Filter(
    ServerPAD,
    'Folder Path' = "/sites/Lightbase-Platform/ServerPAD/Templates/" &&
    IsFolder = false && EndsWith(Name, ".pptx")
  ),
  Name, Ascending
)

```

**Generate button:**

```
Set(
  res,
  GenerateLabelJob.Run(
    txtProductVariantCode.Text,              // sku
    drpLang.Selected.Value,
    Value(txtCount.Text),
    ddTemplate.Selected.Name,                // e.g., label_4x6in_v2.pptx
    JSON({
      TXT_PRODUCT_NAME: txtProductName.Text,
      TXT_CLIENT_ADDRESS: txtClientAddress.Text
    }, JSONFormat.Compact)
  )
);
If(res.status="created", Notify("Job queued", Success), Notify("Failed", Error));

```

### F) Output naming & idempotency

* Output: `Output/<job_id>/<job_id>_<sku>_<yyyymmdd-HHmmss>.pdf`
* Status: `Output/<job_id>/<job_id>.status.json`
* Idempotency: If a job with same `job_id` was processed, either skip or create a deterministic guard (e.g., check for existing status.json).

### G) Designer guidance (quick)

* **Text:** type `{{product_name}}`, `{{client_address}}` → profiler renames to `TXT_*`.
* **Images:** place a sample picture and set **Alt Text** to `{{img:logo}}` or `{{img:qr}}` → profiler renames to `IMG_*`.
* Avoid grouping; ensure shapes have enough area for content; use real label dimensions; embed corporate fonts.

---

## Milestones

**M0 – Environment (1d)**: Folders, `SERVERPAD` var, LabelGen.ps1 + QRCoder.dll.
**M1 – Templates (1–2d)**: Build two templates; verify text/QR fill.
**M2 – PAC (1d)**: Form → Job JSON in `Jobs/incoming`.
**M3 – PAD (1–2d)**: Poll, atomic move, run PS, route done/failed.
**M4 – E2E MVP (1d)**: 1×, 20×, all-fields scenarios.
**M5 – Quality (1d)**: DPI/embedding/retry/error handling.
**M6 – Security (0.5d)**: Service account, least privilege, log policy.
**M7 – Deploy & Train (0.5d)**: Agent rollout, 1-page quick guide.

---

## Gathering Results

**Acceptance**

* All placeholders filled; `TXT_ITEM_NO` padded (e.g., `0001/0060`).
* QR equals `SKU|SERIAL`; scanned OK (3 samples).
* PDF has correct page size; fonts embedded; ≥300 DPI.
* Outputs written to `Output/<job_id>/`; job JSON moved to `done/`.

**KPIs**

* SLA: ≤ **20s** for 10 labels; ≤ **90s** for 100 labels.
* Throughput: ≥ 5 jobs/min (with tuned concurrency).
* Reliability: ≥ 99% success in a 1-week pilot.
* Idempotency: reprocessing same `job_id` doesn’t duplicate outputs.

**Tests**

1. Single label (count=1), all fields.
2. 60× labels: padding, unique QR per slide.
3. Concurrency: 5 jobs × 20 labels.
4. Error cases: missing template / wrong shape name / QR fail → `failed/` with reason.
5. Fonts: verify embed & rendering on another machine.

**Monitoring & Retention**

* `Logs/agent.log` (+ optional per-job logs).
* Retention: **Output 90d**, **Jobs 30d**, **Logs 14d** (configurable).

**Migration to B/Central**

* Add `AllocateRange` Flow & SharePoint List; set `serialProvider.source="central"`, keep QR/template unchanged.

---
