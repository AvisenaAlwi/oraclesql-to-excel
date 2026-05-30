# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [2.0.3] - 2026-05-30

### Changed

- **`archiver` upgraded from `^5.3.2` to `^7.0.1`** — removes transitive `glob@7` (security warnings) and `inflight@1.0.6` (memory-leak warning) from the dependency tree. API is fully compatible; no changes required in consuming code.
- **Backpressure RSS wait: early exit + shorter max timeout** — RSS polling now exits early (~1 s) when RSS is not improving (typical behind a buffering proxy where Buffer pool retention prevents RSS from dropping). Max safety-net timeout reduced from 10 s → 3 s for cases where RSS is actively dropping (V8 GC). In proxy-buffering deployments this eliminates the ~10 s stall per batch that previously wasted time with no memory benefit.

---

## [2.0.2] - 2026-05-29

### Fixed

- **Client disconnect not detected in `pipe()` + `.asZip()` path** — when the HTTP connection was dropped mid-export (browser cancel, ingress timeout), the Oracle fetch loop continued until the next RSS timeout expired. Fixed by adding `close`/`error` listeners on the writable stream in both `pipe()` and `_executeAsZip()`. Export now stops within one batch of disconnect.
- **Backpressure wait timeout reduced 30 s → 10 s** — GC typically recovers in 1–5 s; 30 s was unnecessarily long.

---

## [2.0.1] - 2026-05-29

### Fixed

- **`.compress()` JSDoc** — second parameter `level` (zlib `0`–`9`, default `1`) was undocumented. Now documented with examples.
- **`.compress(bool, level?)` signature** — added `level` parameter to control zlib compression level for both XLSX content and the outer ZIP archive (when `.asZip()` is used). Previously `compress(true)` used ExcelJS default level; now explicitly defaults to `1` (fastest with compression). `compress(false)` remains unaffected (always `level: 0`, store mode).
- **`backpressureThreshold` JSDoc** — documented default was `536870912` (512 MB) but actual default since v2.0.0 is `268435456` (256 MB). Corrected in JSDoc and README.

---

## [2.0.0] - 2026-05-29

### Added

- **`.asZip()`** — new builder method that enables ZIP output mode for `.file()` exports.
  When set, all three terminal methods deliver a single ZIP archive instead of individual `.xlsx` files:
  - `.pipe(res)` — streams the ZIP directly to any Writable (e.g. Express response). Set `Content-Type: application/zip` and `Content-Disposition: attachment; filename="export.zip"` before piping.
  - `.run()` — writes a single `<filePrefix>.zip` file to `outputDir` and returns `ZipRunResult`.
  - `.toBuffer()` — returns the entire ZIP as a `Buffer`. **Not recommended for large data** — holds full ZIP in RAM.
  
  Has no effect when `.file()` is not used (plain `.sheet()` exports are unaffected).

- **`ZipRunResult`** interface exported for TypeScript callers. Contains `file: string` (absolute path to the `.zip`), `success`, `sheets`, `skippedRows`.

### Fixed

- **`showTotalRows` incorrect row range on files 2+ in multi-file exports.** The row range text (e.g. "Showing rows X – Y of Z total") displayed wrong values starting from file 2 because `sheetIndex` reset to `0` for each new file without accounting for previously written rows. Now correctly tracks the global row offset across files.

  *Example fix:* File `Monitoring_1000001-1500000.xlsx` previously showed "Showing rows 1 – 1,000,000 of 2,700,501 total". Now correctly shows "Showing rows 1,000,001 – 1,500,000 of 2,700,501 total".

- **`.pipe()` + `.file()` error message** now explicitly mentions `.asZip()` and the required HTTP headers, making the fix self-evident.

- **`.toBuffer()` + `.file()` error message** updated to mention `.asZip()`.

### Changed

- **`backpressureThreshold` default reduced from 512 MB to 256 MB.** The RSS polling mechanism that pauses Oracle fetching when memory pressure is high now fires earlier, reducing peak RSS during large exports. Users who set `.backpressureThreshold(n)` manually are unaffected.

### Breaking Changes

1. **`backpressureThreshold` default changed.** Exports that previously relied on the 512 MB default will now pause more aggressively at 256 MB. Tune with `.backpressureThreshold(512 * 1024 * 1024)` to restore the old behaviour.

2. **`showTotalRows` display text corrected.** If your code asserts on the exact text of doc-header rows in XLSX output from multi-file exports, update your assertions to reflect the correct row ranges.

3. **`.run()` return type is now `RunResult | MultiRunResult | ZipRunResult`** (previously `RunResult | MultiRunResult`). TypeScript callers that narrowed the return type may need adjustment.

---

## [1.3.0] - 2026-05-29

### Added

- **`OracleSqlToCsv()`** — new builder for streaming Oracle SQL results directly to `.csv`. Uses the same fluent API pattern as `OracleSqlToExcel()`.

  Unlike the Excel path, CSV writes each row directly to the output stream with no intermediate archiver or ZIP buffer. Memory usage is `O(fetchSize × row_size)` at all times regardless of total row count — no backpressure issues, no Ingress/proxy buffering problems, suitable for any data size.

  ```js
  import { OracleSqlToCsv } from '@avisenaalwi/oraclesql-to-excel';

  // HTTP streaming (no memory issues even for 10M+ rows)
  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', 'attachment; filename="report.csv"');
  await OracleSqlToCsv()
    .connectionFactory(() => pool.getConnection())
    .sql('SELECT CODE, NAME, AMOUNT FROM BIG_TABLE')
    .columns([
      { key: 'CODE',   header: 'Code'   },
      { key: 'NAME',   header: 'Name'   },
      { key: 'AMOUNT', header: 'Amount' },
    ])
    .pipe(res);

  // Write to file
  const { file, rowsWritten } = await OracleSqlToCsv()
    .connectionFactory(() => pool.getConnection())
    .sql('SELECT * FROM BIG_TABLE')
    .run('/tmp/export.csv');
  ```

  **Methods:** `.connectionFactory()`, `.sql()`, `.columns()`, `.fetchSize()`, `.separator()`, `.withBom()`, `.onProgress()`.  
  **Terminal:** `.run(filepath)` → `CsvRunResult` · `.pipe(stream)` → `CsvResult` · `.toBuffer()` → `CsvBufferResult`.

- **`CsvResult`**, **`CsvRunResult`**, **`CsvBufferResult`** types exported for TypeScript callers.

---

## [1.2.0] - 2026-05-29

### Added

- **`.file(name, fn)` — multi-file API.** Groups one or more sheets into a named logical file. Each `.file()` call on the builder defines one output `.xlsx` (or a set of split files). This is the primary way to use multi-file exports — it replaces the old builder-level `maxRowsPerFile()` approach.

  ```js
  OracleSqlToExcel()
    .connectionFactory(() => pool.getConnection())
    .outputDir('/tmp')
    .file('laporan', f => f
      .maxRowsPerFile(1_000_000)
      .sheet('Detail',  s => s.sql(SQL1).columns(COLS1).maxRowsPerSheet(900_000))
      .sheet('Summary', s => s.sql(SQL2).columns(COLS2))
    )
    .run()
  ```

  Only supported with `.run()`. Using `.file()` with `.pipe()` or `.toBuffer()` throws an error.

- **`FileConfig`** — new per-file configuration class, obtained via the `.file()` callback. Methods: `.maxRowsPerFile(n)`, `.sheet(name, fn)` (identical to the builder's `.sheet()`).

- **`FileConfig.maxRowsPerFile(n)`** — when set on a `FileConfig`, splits the file into multiple physical `.xlsx` files whenever data rows exceed `n`. Files are named `<name>_<startRow>-<endRow>.xlsx` for single-sheet configs, or `<name>_1.xlsx`, `<name>_2.xlsx`, … for multi-sheet configs. Oracle's `ResultSet` is kept open across files — only one query and one connection per sheet regardless of how many files are produced. Sheet splitting (`.maxRowsPerSheet()`) applies independently within each file.

- **`MultiRunResult`** — return type of `.run()` when `.file()` is used. Contains a `files` array of `FileSegment` (`{ file, startRow, endRow }`).

- **`FileSegment`** type exported for TypeScript callers.

- **Multiple `.file()` calls** — each call defines an independent output file. All files are written sequentially and results are merged into a single `MultiRunResult`.

### Backward compatibility

- **`.sheet()` at builder level is unchanged.** Existing single-file exports require no migration.
- **`.pipe()` and `.toBuffer()` are unchanged** when `.file()` is not used.

---

## [1.1.4] - 2026-05-28

### Fixed

- **`.run()` had no RSS throttling** — the RSS-based polling introduced in `v1.1.3` only applied to `.pipe()`. When using `.run()` (write to file), `drainFn` was always `null` and memory grew unchecked. On slow container storage (OpenShift overlay filesystem, network PVC) the archiver buffer accumulates data at the same rate as the `.pipe()` + reverse proxy case.

  RSS polling is now also applied for `.run()`: when `process.memoryUsage().rss` exceeds the configured threshold, the Oracle fetch pauses until disk writes drain the buffer and RSS drops.

---

## [1.1.3] - 2026-05-28

### Fixed

- **Memory still grew linearly behind a reverse proxy** — when Node.js is behind nginx or any reverse proxy, the proxy accepts data from Node.js instantly (local socket), so `write()` on the output stream always returns `true`. The event-driven drain loop introduced in `v1.1.2` never triggered. Data accumulated in Node.js's internal `outputData` buffer (visible as RSS growth) while the proxy slowly forwarded data to the slow end client.

  Added a second, RSS-based fallback: after each Oracle fetch batch the library checks `process.memoryUsage().rss`. If it exceeds the configured threshold, it polls every 200 ms and waits until RSS drops before fetching the next batch. This approach works regardless of proxy topology because it monitors the process's own memory, not stream events.

### Changed

- **`.backpressureThreshold(bytes)` is no longer a no-op** — repurposed as the RSS threshold for the polling fallback. Default changed from `16 MB` to `512 MB`. Set lower if your PM2 `--max-memory-restart` limit is below 1 GB.

---

## [1.1.2] - 2026-05-28

### Fixed

- **Archiver buffer accumulation not fully drained between batches** — `v1.1.1` waited for one `drain` event on the output stream before fetching the next Oracle batch. However, after a TCP `drain`, the ExcelJS internal archiver immediately flushes its own queued data back into the stream, re-filling the buffer. This cycle meant only one TCP buffer's worth of data (~4–8 MB) was drained per batch pause, while the archiver continued to accumulate. Over many batches this still caused out-of-memory on slow clients.

  The drain wait is now a loop: after each `drain` event the code yields one event-loop turn (via `setImmediate`) to let the archiver flush pending data, then checks for backpressure again. The loop exits only when no new data arrives in one turn — meaning the archiver is truly empty and the next Oracle fetch can safely begin.

- **`bytesSinceDrain` threshold could trigger false-positive drain waits** — the proactive byte-count check introduced in `v1.1.0` set `needsDrain = true` based on bytes written since last drain, even when `write()` was still returning `true` (no actual backpressure). With the drain loop now in place, this check is unnecessary and was removed. Backpressure detection now relies solely on `write()` returning `false`, which is the authoritative Node.js signal.

### Deprecated

- **`.backpressureThreshold(bytes)`** — now a no-op. Kept for API compatibility with `v1.1.0`. See fix above.

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
