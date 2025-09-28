# LGMS Deployment & Configuration Checklist

✅ Use this checklist to verify that a deployment was completed successfully.

---

## 🧰 Environment

- [ ] Windows 10/11 device
- [ ] PowerShell 5.1 installed
- [ ] PowerPoint (Desktop) installed
- [ ] OneDrive for Business syncing active

---

## 📁 Folder Structure

- [ ] `Templates/Raw/` and `Templates/Profiled/`
- [ ] `Jobs/incoming/`, `processing/`, `done/`, `failed/`
- [ ] `Output/` and `Logs/`
- [ ] `Tools/LabelGen/`, `Libs/`, `Setup/`, `TemplateProfiler/`

---

## 📦 Required Files

- [ ] `LGMS.msapp` imported into Power Apps
- [ ] `LabelGen.ps1`, `TemplateProfiler.ps1` present
- [ ] `ZXing.Net.dll`, `QRCoder.dll` present in `Tools/Libs/`

---

## ⚙️ Setup Actions

- [ ] App imported and file connections configured
- [ ] Agent installed via `deploy-ServerPAD.cmd`
- [ ] Critical files pinned using `attrib +P`

---

## ✅ Functional Test

- [ ] `.profile.json` auto-generated from `.pptx`
- [ ] Job JSON created by Power App
- [ ] PDF output created in correct folder
- [ ] Logs written to correct date log file
