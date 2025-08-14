# Runbook — PAD Agent (LabelGen)

## Scope
Windows agent that processes jobs from SharePoint-synced folder `ServerPAD/Jobs`.

## Prereqs
- PowerPoint installed; PAD installed
- Env var `SERVERPAD = "C:\Users\Learner\LightBase\Lightbase-Platform - ServerPAD"`
- `Tools/LabelGen/LabelGen.ps1` and `QRCoder.dll` present

## Start/Stop
- Start PAD and run flow **LabelGen-Worker**
- To restart: stop PAD, kill any stuck POWERPNT.EXE, start PAD

## Config
- Poll interval: 5s
- Folders: `Jobs/incoming|processing|done|failed`, `Output/<job_id>/`

## Flow Outline
1. Move `incoming/*.json` → `processing/<job_id>.lock.json`
2. Run:
powershell.exe -ExecutionPolicy Bypass ^
-File "%SERVERPAD%\Tools\LabelGen\LabelGen.ps1" ^
-JobJsonPath "%jobMovedPath%" -Root "%SERVERPAD%"
3. On exit code 0 → move job JSON to `done/` else `failed/`

## Health Checks
- New job moves to `processing/` < 10s
- `Logs/agent.log` gets updated
- `Output/<job_id>/thumb.png` exists

## Common Errors & Recovery
- **Template missing** → job to `failed/`; fix template path; move JSON back to `incoming/`
- **COM error** → restart PowerPoint & PAD
- **Permissions/Sync** → check OneDrive status and library permissions
