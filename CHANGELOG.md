# Changelog

## v1.0.0 - 2026-05-23

Filings Atlas / 全球披露图谱 v1.0.0 is the first GitHub release for the renamed product. It keeps the project PDF-only and ships the 7-market Windows desktop app.

### Added

- Product rename from **FS Capture** to **Filings Atlas / 全球披露图谱**.
- Runtime Chinese / English UI switching with persistent language setting.
- Taiwan market support through TWSE + MOPS.
- Japan market support through EDINET.
  - EDINET Subscription-Key is strongly recommended for v1.0.
  - Public fallback is retained, but the official EDINET API requires `Subscription-Key`.
- United Kingdom market support through FCA National Storage Mechanism.
  - UK does not require an API key.
- Korea no-key mode through DART public disclosure pages, with DART API key kept as an optional accelerator.
- Incremental update mode based on sidecar metadata.
- Sidecar metadata storage under `data/cache/sidecars/`.
- "How to add a new market" extension guide in `ARCHITECTURE.md`.
- GitHub Actions release workflow for `FilingsAtlas-v1.0.0-windows.zip`.

### Changed

- Output filenames now use the stable flat PDF contract:
  `{exchange}_{code}_{company}_{year}_{kind_zh}.pdf`.
- IPO/prospectus output path helpers are unified across supported markets.
- UI strings are centralized in bilingual dictionaries.
- Playwright rendering uses a process-wide browser pool for HTML-to-PDF filings.
- Large PDF downloads support `.part` resume and retry.
- Name resolver caches use single-flight loading to reduce duplicated remote requests.

### Fixed

- HK annual report selection uses PDF verification before falling back to title/date heuristics.
- HK non-December fiscal year mappings are expanded.
- A-share akshare code/name dirty data is diagnosed with strict zip handling and fallback.
- KR industry metadata reads the official `induty_code` field.
- Batch import reports rejected rows without treating company names as errors.
- SEC paged submissions fallback is covered for older filings.
- TWSE/MOPS Big5, ROC-year and temporary PDF URL handling are covered by tests.

### Breaking Changes

- The visible product name and executable are now `Filings Atlas.exe`.
- GitHub release assets use `FilingsAtlas-v1.0.0-windows.zip`.
- The repository directory may still be named `FS Capture`; this is intentionally not renamed.
- Sidecar JSON files are migrated away from `output/` to `data/cache/sidecars/`.

### Validation

- Non-e2e test matrix: `145 passed, 7 deselected`.
- Ruff lint: `All checks passed`.
- UK smoke verified real `%PDF` outputs for `ULVR`, `HSBA`, and `AZN`.
- JP official API smoke without key returns invalid subscription key; configure EDINET Subscription-Key for reliable JP downloads.
