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
