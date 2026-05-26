# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.1.0] - 2026-05-26

### Fixed

- **Backpressure not respected during `.pipe()` streaming** — when using `.pipe(res)` with a slow HTTP client, Oracle rows were fetched and committed to the ExcelJS stream faster than the client could receive them. This caused the Node.js stream internal buffers to grow unboundedly, leading to process memory exhaustion on large exports (1M+ rows).

  The library now checks the output stream's `writableLength` after each Oracle fetch batch. If it exceeds the threshold, the fetch is paused until the stream drains before continuing.

### Added

- **`.backpressureThreshold(bytes)`** — configurable byte threshold that controls when the Oracle fetch is paused to allow the output stream to drain. Default: `16777216` (16 MB). Only applies to `.pipe()`; has no effect on `.run()` or `.toBuffer()`.

---

## [1.0.0] - 2026-05-01

### Added

- Initial release.
- Streaming Oracle SQL → Excel (`.xlsx`) via `ExcelJS` streaming writer.
- Fluent builder API: `.connectionFactory()`, `.sheet()`, `.run()`, `.pipe()`, `.toBuffer()`.
- Per-sheet: `.sql()`, `.columns()`, `.fetchSize()`, `.maxRowsPerSheet()`, `.freezeHeader()`, `.autoFilter()`, `.headerStyle()`, `.docHeader()`, `.showTotalRows()`, `.onRowError()`.
- Multi-sheet and automatic split sheets.
- Document header rows with simple and multi-column (merge) modes.
- Row range summary via parallel `COUNT(*)` query.
- Progress callback (`.onProgress()`).
- ZIP compression toggle (`.compress()`).
- Auto-detect columns from Oracle metadata when `.columns()` is omitted.
