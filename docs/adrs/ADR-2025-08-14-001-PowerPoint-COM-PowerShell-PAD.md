# ADR-2025-08-14-001 — Use PowerPoint COM + PowerShell + PAD

## Context
Need designer-owned layout, corporate licensing compatibility, fast offline generation, SharePoint integration.

## Decision
Use **PowerPoint COM** for rendering, **PowerShell (LabelGen.ps1)** for automation & QR, orchestrated by **Power Automate Desktop (PAD)**.

## Consequences
**Pros:** No new licenses; designers free to move shapes; fast local PDF; simple ops  
**Cons:** Windows-only agent; COM error handling; needs Office installed

## Alternatives Considered
- HTML+wkhtmltopdf (layout fidelity risk)
- Office Open XML SDK only (complex image handling)
- LibreOffice headless (rendering differences)

## Status
Accepted — 2025-08-14
