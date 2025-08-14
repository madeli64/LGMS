# Runbook — PAC Flow (from Power Apps)

## Trigger
Power Apps button with params:
- `product_variant_code` (sku)
- `language` (en-CA|fr-CA)
- `count` (int)
- `templateFilename` (e.g., `label_4x6in_v2.pptx`)
- `fieldsJson` (JSON string)

## Steps (high-level)
1) `Compose` Job JSON as per spec (envelope + fields/images)
2) `Create file` → SharePoint: `ServerPAD/Jobs/incoming/<job_id>.json`
3) Respond to Power Apps: `{ status: 'created', jobUrl: '...' }`

## Notes
- Validate template filename ends with `.pptx`
- Use `guid()` for `job_id`
- Language set: `en-CA` or `fr-CA` only
