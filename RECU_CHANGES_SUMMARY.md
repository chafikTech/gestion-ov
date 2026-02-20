# Reçu (Total) - Summary (DOCX-only)

## Update (February 11, 2026)

The application now generates **DOCX only** (no PDF) via Python (`python-docx`).

## Expected behavior

- Generates **one combined receipt per month** for all workers: `Recu_MM_YYYY.docx`
- The receipt shows the **monthly total** (not one receipt per worker)

## Implementation

- UI trigger: `src/frontend/index.html` + `src/frontend/js/documents.js`
- Backend orchestration: `src/backend/documents.js` (documentType: `recu-combined`)
- Python generator: `src/python/generate_document.py` (writes the `.docx` file)

## Python JSON contract (stdin/stdout)

- Success:
  - `{"success": true, "docxFileName":"...", "docxFilePath":"..."}`
- Error:
  - `{"success": false, "message":"..."}`
