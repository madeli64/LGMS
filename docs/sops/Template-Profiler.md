# SOP — Template Profiler (Designer workflow)

## Goal
Standardize placeholders in PPTX and auto-generate `<template>.profile.json`.

## How designers mark placeholders
- **Text placeholders:** type `{{product_name}}`, `{{client_address}}` in text boxes
- **Image placeholders:** insert a sample image, set **Alt Text** to `{{img:logo}}` or `{{img:qr}}`
- Do **not group** pictures; turn off "Compress Pictures"

## Run Profiler
- Run `Tools/TemplateProfile.ps1 <path-to-template.pptx>`
- The script:
  - Scans **Text + Alt Text** for `{{...}}`
  - Renames shapes to `TXT_*` / `IMG_*` (ensures unique names with `__2`)
  - Creates `<template>.profile.json` listing fields/images and `shapes[]`
  - Defaults: `TXT_ITEM_NO = "{I}/{N}"`, all fields **optional** (required set later in app)

## Output
- `Templates/<name>.pptx`
- `Templates/<name>.profile.json`

## Validation
- Open the template and check shape names in Selection Pane
- Ensure repeated fields show multiple shape names in profile’s `shapes[]`
