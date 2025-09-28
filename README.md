# LGMS — Label Generator Management System

**LGMS (Label Generator Management System)** is a Power Apps + PowerShell solution for fully local label generation using manifest-driven templates and PowerPoint COM automation. It requires no cloud rendering or premium connectors.

---

## 📌 Key Features

- Template gallery with live PDF preview.
- Manifest-based dynamic input form generation.
- PowerPoint COM-based rendering (local-only).
- Auto-numbering (1/N, N/N), QR/Barcode embedding.
- Idempotent job execution (based on GUID).
- File-based job queue (incoming/processing/done/failed).
- Zero reliance on SharePoint Lists or Dataverse.
- Operator-focused UI with minimal dependencies.
- Logging and structured output archive.

---

## 🧱 Technologies Used

- **Frontend**: Power Apps Canvas App (`LGMS.msapp`)
- **Data Storage**: JSON files in SharePoint Document Library synced via OneDrive
- **Local Agent**: PowerShell (`ServerPAD-Agent.ps1`) + Scheduled Task
- **Rendering**: PowerPoint COM, QRCoder.dll, ZXing.Net.dll
- **Template Processing**:
  - `TemplateProfiler.ps1`: creates `.profile.json` from `.pptx`
  - `LabelGen.ps1`: renders data into PDF

More technical details in [`docs/specs/ENGINE.md`](docs/specs/ENGINE.md)

---

## 🚀 Deployment Instructions

### 1. Prerequisites

- Windows 10/11 with PowerShell 5.1
- Microsoft PowerPoint (Desktop)
- OneDrive for Business account synced to SharePoint Library
- Access: Owner/Editor on target Document Library

Test COM:

```powershell
$pp = New-Object -ComObject PowerPoint.Application; $v=$pp.Version; $pp.Quit(); "OK v$($v)"
```

---

### 2. Folder Structure

Create and sync a SharePoint Document Library:

```
ServerPAD/
├── Templates/
│   ├── Raw/
│   └── Profiled/
├── Jobs/
│   ├── incoming/
│   ├── processing/
│   ├── done/
│   └── failed/
├── Output/
├── Tools/
│   ├── LabelGen/
│   ├── Libs/
│   ├── Setup/
│   └── TemplateProfiler/
└── Logs/
```

---

### 3. Import Power Apps

1. Go to [make.powerapps.com](https://make.powerapps.com)
2. Import `LGMS.msapp`
3. Configure file connections to your synced library
4. Publish and Share the app

---

### 4. Install Local Agent

Run:

```bat
cd /d "%ROOT%\Tools\Setup"
deploy-ServerPAD.cmd
```

Creates:

- `C:\SPAD\run-agent.vbs`
- Scheduled Task: **ServerPAD Agent** (runs on logon)
- Immediate run + log output

---

### 5. Sync + Pin Key Files

Prevent OneDrive from offloading:

```bat
attrib +P -U "%ROOT%\Templates\Raw\*.pptx"
attrib +P -U "%ROOT%\Templates\Profiled\*.profile.json"
attrib +P -U "%ROOT%\Tools\**\*.*"
```

---

### 6. Sanity Check

1. Place a `.pptx` in `Templates\Raw` → Generates `.profile.json`
2. Create a Job in App → `job.json` created
3. PDF appears in `Output/<Template>/` after agent runs
4. Check logs in `Logs\agent.YYYY-MM-DD.log`

---

### 7. Troubleshooting

- COM error `80080005`? Ensure Scheduled Task runs in user session with delay.
- Agent not running? Check:
  ```bat
  schtasks /query /tn "ServerPAD Agent"
  reg query HKCU\...Run
  ```

---

## 📂 Project Structure

```
.
├── LGMS.msapp
├── Tools/
│   ├── LabelGen/
│   ├── Libs/
│   ├── Setup/
│   └── TemplateProfiler/
├── docs/
│   └── specs/
│       └── ENGINE.md
│   └── runbooks/
│       ├── INSTALL.md
│       └── CHECKLIST.md
├── LICENSE
└── README.md
```

---

## 📘 See Also

- [docs/specs/ENGINE.md](docs/specs/ENGINE.md) — Script internals, JSON format, and rendering flow
