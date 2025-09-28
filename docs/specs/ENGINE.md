# Label Engine — Technical Overview

This document describes the two core scripts powering the LGMS rendering system:

- `TemplateProfiler.ps1`: Extracts field definitions from PowerPoint templates.
- `LabelGen.ps1`: Renders final PDFs with data injection, barcode/QR embedding, and output saving.

---

## 🧩 TemplateProfiler.ps1

### 📌 Purpose

Parses a `.pptx` template and generates a manifest JSON (`*.profile.json`) used by the Power Apps frontend to dynamically render input forms.

### 🛠️ Input

- **PowerPoint File**: Must contain placeholder tags in the format `{{FIELD}}` in text boxes.

### ⚙️ Behavior

1. Opens the `.pptx` file using **PowerPoint COM**.
2. Iterates through all slides and shapes.
3. Extracts unique placeholder names (`{{FIELD}}`, `{{I}}`, `{{N}}`).
4. Outputs a JSON structure listing fields.

### 📝 Output

Creates a manifest file (`*.profile.json`) with a structure like:

```json
{
  "template": "ProductLabel_v2",
  "fields": [
    { "name": "Name", "type": "string", "required": true },
    { "name": "Batch", "type": "string", "required": true }
  ],
  "special": ["I", "N"]
}
```

### 📍 Output Location

Should be saved under:

```
Templates/Profiled/
```

---

## 🧩 LabelGen.ps1

### 📌 Purpose

Executes the final PDF generation for a given job using the `.pptx` template and the associated `.profile.json`.

### 🛠️ Input

- **Job Definition** (`job.json`) — Contains:
  - `template`: name of the template
  - `manifest`: path to `.profile.json`
  - `data`: an array of records (field values)
  - `count`: number of labels
  - `outputFile`: destination for final PDF

### ⚙️ Behavior

1. Opens the template `.pptx` via **PowerPoint COM**.
2. Duplicates slides as needed per record.
3. Replaces placeholders (`{{FIELD}}`, `{{I}}`, `{{N}}`) with actual data.
4. Generates **1/N** and **N/N** counters.
5. Embeds **Barcode (Code128)** and/or **QR Code** using:
   - `ZXing.Net.dll`
   - `QRCoder.dll`
6. Saves the final document as a **PDF**.

### 📝 Output

- Single PDF file per job
- Saved to `Output/<TemplateName>/`
- Example: `Output/ProductLabel_v2/Job-XYZ123.pdf`

### 📂 Dependencies

Must exist in:

```
Tools/Libs/
├── ZXing.Net.dll
└── QRCoder.dll
```

---

## ⚠️ Error Handling

- **If COM or Office is unavailable**: script exits with error.
- **If placeholders are unmatched**: may result in blank fields or skipped inserts.
- **On failure**: Job is moved to `Jobs/failed/` and detailed error is logged.

---

## 🧪 Test Execution (Manual)

Test TemplateProfiler:

```powershell
powershell -ExecutionPolicy Bypass -File "Tools/TemplateProfiler/TemplateProfiler.ps1" -Input "Templates/Raw/ProductLabel.pptx" -Output "Templates/Profiled/ProductLabel.profile.json"
```

Test LabelGen:

```powershell
powershell -ExecutionPolicy Bypass -File "Tools/LabelGen/LabelGen.ps1" -Job "Jobs/incoming/abc123/job.json"
```

---

## ✅ Summary

| Script               | Purpose                      | Input           | Output                |
|---------------------|------------------------------|------------------|------------------------|
| TemplateProfiler.ps1| Extract field manifest        | `.pptx` file     | `.profile.json`       |
| LabelGen.ps1        | Generate final PDF with data  | `job.json`       | `Output/<template>/`  |
