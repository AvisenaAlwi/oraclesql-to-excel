# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.1.1] - 2026-05-26

### Fixed

- **Backpressure check was using the wrong signal** — `v1.1.0` checked `writableLength` on the output stream (`res`), but data actually accumulates in ExcelJS's internal archiver buffer which sits upstream of `res`. This caused the check to never trigger and memory still grew unboundedly on slow clients.

  The fix intercepts `write()` on the output stream directly. When `write()` returns `false` — Node.js's authoritative backpressure signal — the Oracle fetch is paused until the stream drains.

- **Client disconnect leaked Oracle connection for up to 30 seconds** — when a browser cancelled a download mid-stream, the `drain` event never fired, and the export continued fetching from Oracle until the 30-second safety timeout. The stream `close` and `error` events now abort the fetch immediately, closing the Oracle connection and result set right away.

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
