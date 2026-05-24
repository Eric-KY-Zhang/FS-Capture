# Changelog

## v1.0.0 - 2026-05-24

First GitHub release for Filings Atlas / 全球披露图谱. This release adds Singapore SGX support, improves batch fetch performance, and keeps the PDF-only scope unchanged.

### Added

- Singapore market support through SGXNet public disclosures.
  - Annual reports verified with DBS (`D05`), UOB (`U11`) and Singtel (`Z74`).
  - Interim report support verified with UOB H1.
  - IPO prospectus support verified with `3407` / LION-CM EM ASIA INDEX ETF.
- Japan EDINET public web fallback fully implemented (`edinet_web.py`).
  - JP downloads now work without an EDINET API key out of the box.
  - API key remains an optional accelerator with higher rate limit (`edinet = 2.0` vs `edinet_web = 1.0`).
  - Public mode tested with Toyota (`7203`), Sony (`6758`) and SoftBank (`9984`) 2024 annual reports.
- Benchmark harness under `development/tests/benchmarks/` with explicit `FS_CAPTURE_RUN_BENCHMARK=1` opt-in.
- Performance reports under `docs/perf/` for v0.9 baseline, v1.0 batch-5 A/B results and bundle size.

### Changed

- Default worker concurrency increased from 4 to 6.
- Name resolver warm-up now starts in the background for full-map markets.
- Large PDF streaming chunk size increased from 64 KB to 256 KB.
- Playwright rendering now keeps thread-local browser/context state and clears cookies/permissions after each render.
- HTTP clients now use tuned connection limits.
- Sidecar writes are atomic temp-file replacements.
- PyInstaller spec includes the SG plugin and extra excludes for unused test/dev stacks.
- Build script trims non-Chinese/non-English Qt translation files after packaging.

### Performance

- Batch benchmark improved from 211.76s to 94.22s on the same 21-task A/HK/US/KR/TW/UK/SG set, a 55.5% reduction versus v0.9 baseline.
- The prior UK `AZN` zero-report failure in concurrent Playwright rendering is resolved.
- Bundle size changed from 449.0 MB to 441.4 MB (-7.6 MB). Higher-risk Playwright selective collection and UPX compression were intentionally skipped.

### Compatibility

- v0.9 sidecars under `data/cache/sidecars/` remain readable.
- The flat output naming contract remains unchanged: `{exchange}_{code}_{company}_{year}_{kind_zh}.pdf`.
- Existing `config.toml` fields remain compatible.

## v0.9.0 - 2026-05-23

Filings Atlas / 全球披露图谱 v0.9.0 is the internal iteration for the renamed product. It keeps the project PDF-only and ships the 7-market Windows desktop app.

### Added

- Product rename from **FS Capture** to **Filings Atlas / 全球披露图谱**.
- Runtime Chinese / English UI switching with persistent language setting.
- Taiwan market support through TWSE + MOPS.
- Japan market support through EDINET API mode.
  - The no-key public crawler was completed later in the v1.0 addendum.
- United Kingdom market support through FCA National Storage Mechanism.
  - UK does not require an API key.
- Korea no-key mode through DART public disclosure pages, with DART API key kept as an optional accelerator.
- Incremental update mode based on sidecar metadata.
- Sidecar metadata storage under `data/cache/sidecars/`.
- "How to add a new market" extension guide in `ARCHITECTURE.md`.
- GitHub Actions release workflow preparation.

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
- Public release assets are deferred to the next release.
- The repository directory may still be named `FS Capture`; this is intentionally not renamed.
- Sidecar JSON files are migrated away from `output/` to `data/cache/sidecars/`.

### Validation

- Non-e2e test matrix: `145 passed, 7 deselected`.
- Ruff lint: `All checks passed`.
- UK smoke verified real `%PDF` outputs for `ULVR`, `HSBA`, and `AZN`.
- JP API-key path had unit coverage; public no-key crawler work was deferred to v1.0.
