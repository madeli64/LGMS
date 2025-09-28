# LGMS Installation Guide

This guide walks through a complete setup of the LGMS (Label Generator Management System), including OneDrive sync, template setup, agent deployment, and Power Apps configuration.

---

## 📁 1. Create Document Library Structure

In your SharePoint site:

- Create a Document Library named `ServerPAD`
- Sync it to your local machine via OneDrive

Structure:

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
├── Logs/
└── Tools/
    ├── Setup/
    ├── LabelGen/
    ├── Libs/
    └── TemplateProfiler/
```

---

## 💾 2. Pin Critical Files (Always Available)

Run:

```bat
attrib +P -U "%ROOT%\Templates\Raw\*.pptx"
attrib +P -U "%ROOT%\Templates\Profiled\*.profile.json"
attrib +P -U "%ROOT%\Tools\**\*.*"
```

This prevents OneDrive from offloading important scripts/templates.

---

## 🧠 3. Install Microsoft PowerPoint

Required for COM-based rendering.

Test COM access:

```powershell
$pp = New-Object -ComObject PowerPoint.Application; $v=$pp.Version; $pp.Quit(); "OK v$($v)"
```

---

## ⚙️ 4. Install LGMS App

1. Go to [make.powerapps.com](https://make.powerapps.com)
2. Import `LGMS.msapp`
3. Update connections to match your synced SharePoint Library
4. Set constant paths as needed
5. Publish and Share

---

## ⚡ 5. Install Local Agent

Open CMD as **Administrator**:

```bat
cd /d "%ROOT%\Tools\Setup"
deploy-ServerPAD.cmd
```

Creates:

- `C:\SPAD\run-agent.vbs`
- Scheduled Task: `ServerPAD Agent` (ONLOGON)
- Immediate run + log output

Alternate (PowerShell):

```powershell
.\ServerPAD-Agent.ps1 -Install -Root "C:\Users\...\ServerPAD"
```

---

## 🔍 6. Test End-to-End

1. Place a `.pptx` in `Templates/Raw` → Should auto-generate `.profile.json`
2. Create a job in the Power App
3. Agent will process → output PDF in `Output/Template/`
4. Check logs in `Logs\`

---

## 🛠 7. Uninstall / Reset

From:

```bat
cd /d "%ROOT%\Tools\Setup"
Reset-ServerPAD.cmd
uninstall-ServerPAD.cmd
```
