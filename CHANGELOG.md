# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [2.4.0] - 2026-07-01

### Added

- **`.transform(fn)` on `SheetConfig` (Excel)** — per-row value transform. Receives the assembled row after all per-column transforms and must return a new row object. Supports async (returns `Promise`). A dev-only warning is logged once per transform key when a `Promise` is detected (`NODE_ENV` not `production`/`prod`).
- **`.transform(fn)` on `OracleSqlToCsvBuilder`** — same per-row transform API as Excel, applied in the CSV write path.
- **`ColumnDef.transform`** — per-column value transform function: `(value, rawRow) => unknown`. For Excel, `value` is the post-`castCell` typed value. For CSV, `value` is the raw Oracle value. The return value is written directly — no re-casting. Supported on both Excel and CSV builders.

---

## [2.3.2] - 2026-06-04

### Fixed

- **`headerGroups` phantom-row bug with `mergeDown`** — `ws.mergeCells()` in ExcelJS streaming internally reserves the target row, causing the next `ws.addRow([])` to skip that row number. Sub-column labels were being written to the wrong row, leaving an empty phantom row between group headers and sub-headers. Fixed by using `ws.getRow(expectedRowNum)` for rows after the first instead of `ws.addRow([])`.
- **`mergeDown` cells not rendering as merged** — blocked cells (those covered by a vertical merge from a previous row) were never written to the worksheet XML, so ExcelJS did not apply the merge visually. Fixed by explicitly writing an empty value to each blocked cell so it appears in the row XML.
- **Column blocking used `row.number` from ExcelJS** — due to the phantom-row issue, `row.number` for subsequent rows was higher than expected, breaking the `blockedUntil` check. Fixed by using a relative internal counter (`relRow`) instead of `row.number` for all blocking comparisons.

---

## [2.3.1] - 2026-06-04

### Added

- **`align` on `HeaderStyle`** — controls horizontal text alignment of the column header row. Accepts `'left'`, `'center'`, or `'right'`. Example: `.headerStyle({ bgColor: '4472C4', fontColor: 'FFFFFF', bold: true, align: 'center' })`.

### Fixed

- **`headerGroups` rendered an extra duplicate header row** — when `.headerGroups()` was used, the last group row (sub-column labels) and the regular column header row were both written, producing a duplicate row with filter dropdowns on the wrong row. The fix: when `headerGroups` is provided the last group row serves as the column header row and `writeHeaderRow` is skipped. `autoFilter` and `freezeHeader` now target the last group row correctly.

---

## [2.3.0] - 2026-06-04

### Added

- **`.headerGroups(rows)` on `SheetConfig`** — multi-level grouped header rows written immediately above the column header row. Accepts the same `DocHeaderRow[]` format as `.docHeader()`, with full support for `mergeAcross` (additional columns to the right) and `mergeDown` (additional rows downward). Columns blocked by a `mergeDown` cell from a previous row are skipped automatically. `.freezeHeader()` and `.autoFilter()` continue to target the actual column header row, not the group rows.

---

## [2.2.0] - 2026-06-04

### Added

- **`.locale(value)` on `OracleSqlToExcelBuilder` and `OracleSqlToCsvBuilder`** — BCP 47 locale tag used to format numbers in the `"Showing rows X – Y of Z total"` summary produced by `.showTotalRows()`. Controls the thousand and decimal separators. Default: `'en-US'` (`1,000,000`). Example: `.locale('id-ID')` → `1.000.000`.
- **`.docHeader()` on `OracleSqlToCsvBuilder`** — prepend one or more rows above the column header on every CSV file, matching the API already available on `SheetConfig`. Accepts an array of `DocHeaderRow` objects. Only `text` (simple mode) and `columns[].text` (column mode) are written; style, merge, and height are ignored in plain-text CSV. The doc header is written on every file segment when `.maxRowsPerFile()` is used.

### Changed

- **Builder internal fields and methods are now truly `private`** — `OracleSqlToExcelBuilder` and `OracleSqlToCsvBuilder` internal members (all `_`-prefixed fields and `_execute*` methods) now carry the TypeScript `private` keyword in addition to the existing `/** @private */` JSDoc tag. IDE autocomplete no longer surfaces these members when working with a builder instance.

---

## [2.1.0] - 2026-06-03

### Fixed

- **"Showing rows" end value was wrong when `maxRowsPerFile < maxRowsPerSheet`** — the end row was calculated using `_maxRowsPerSheet` (e.g. the default 1,000,000) instead of being capped by the actual file row limit. For example, with `maxRowsPerFile(500_000)` and default `maxRowsPerSheet`, file 1 incorrectly showed `"Showing rows 1 – 1,000,000 of 2,700,636 total"` instead of `"Showing rows 1 – 500,000 of 2,700,636 total"`. Fixed by also capping end with `globalRowOffset + maxRows` in the `Math.min` calculation.

### Changed

- **Dual "File / Sheet" row range labels when multiple sheets exist per file** — when `maxRowsPerFile > maxRowsPerSheet` (i.e. one file can contain more than one sheet), the row range summary is now split into two lines:
  - **`File: Showing rows X – Y of Z total`** — appears on the **first sheet only**, showing the row range covered by the entire file.
  - **`Sheet: Showing rows X – Y of Z total`** — appears on **every sheet**, showing the row range for that individual sheet.
  - When `maxRowsPerFile ≤ maxRowsPerSheet` (at most one sheet per file), only a single `"Showing rows X – Y of Z total"` line is shown (no prefix), unchanged from before.
- **`_executeSheet` (single-file path) now prefixes `"Sheet: "` on split sheets** — when total rows exceed `maxRowsPerSheet` inside a single file, each sheet's row range summary now carries the `"Sheet: "` prefix to clarify it refers to that sheet's slice, not the full dataset.

---

## [2.0.7] - 2026-06-02

### Fixed

- **`OracleSqlToCsv().pipe()` left the response stream open after export completed** — the writable stream was never ended on the success path, causing HTTP clients to hang indefinitely waiting for more data even though all rows had been written. The error path already called `stream.end()`, but the normal completion path did not. Fixed by calling `stream.end()` after `_execute()` resolves, matching the behaviour of `.run()` and `.toBuffer()`.

---

## [2.0.6] - 2026-06-02

### Fixed

- **`freezeHeader` threw "Cannot set property views of #\<WorksheetWriter\> which has only a getter"** — `ExcelJS.stream.xlsx.WorkbookWriter.addWorksheet()` returns a `WorksheetWriter` whose `views` property is exposed as a getter with no setter. Assigning `worksheet.views = [...]` fails at runtime even though TypeScript allows it (the internal `StreamWorksheet` alias pointed to the non-streaming `Worksheet` type which does have a setter). Fixed by mutating the existing array in place via `splice()` instead of reassigning. Affects both the single-sheet (`_executeSheet`) and multi-segment (`_executeSheetSegment`) code paths whenever `.freezeHeader(true)` is set.

---

## [2.0.5] - 2026-05-30

### Fixed

- **`prevFileNote` ("Previous data on file: …") now appears above "Showing rows…" and the column header** — in v2.0.4 the note was inserted *after* the column header row, placing it between the header and the first data row. It now appears before the "Showing rows X – Y of Z total" summary line and before the column header row, matching the expected visual order at the top of each file's first sheet.

---

## [2.0.4] - 2026-05-30

### Added

- **Cross-file navigation notes** — when a query is split across multiple files (via `maxRowsPerFile`), each file now contains informational rows at the sheet boundary:
  - First sheet of file 2+: `"Previous data on file: {name}_1.xlsx"` (italic, gray) immediately after the column header row.
  - Last active sheet of files 1 to N-1: an empty row followed by `"Next data available on file: {name}_2.xlsx"` (italic, gray) before the file ends.
  - Notes use the same visual style as the existing within-file sheet continuation notes (`"Continued on sheet: …"` / `"Continued from sheet: …"`).

### Fixed

- **Sheet split no longer creates empty sheets at file boundary** — when `rowCounter >= maxRowsPerSheet` occurred on the exact same row as `fileRowsWritten >= maxRowsPerFile`, an empty continuation sheet was created inside the current file before earlyReturn fired. The split check now skips the sheet transition when the file limit is also reached, preventing the spurious empty sheet.
- **`maxRowsPerSheet` no longer interferes with `maxRowsPerFile` enforcement** — previously, if `maxRowsPerSheet` was set larger than `maxRowsPerFile`, the sheet split could trigger at exactly the file boundary, creating an empty sheet that received the cross-file navigation note instead of the last data sheet. This is now handled correctly.

### Changed

- **File naming simplified to `{name}_{part}.xlsx`** — `.run()` without `.asZip()` (`MultiRunResult`) now uses sequential numbering (`data_1.xlsx`, `data_2.xlsx`) instead of row-range naming (`data_1-1000000.xlsx`). ZIP entries (`.asZip()`) were already using sequential naming.
- **`FileSegment` now includes per-sheet row info** — a `sheets: FileSheetInfo[]` array replaces the old flat `startRow`/`endRow` fields. Each entry is `{ name: string, startRow: number, endRow: number }` representing the row range of one `.sheet()` query within the file. Multiple `.sheet()` calls each get their own entry. The old `FileSegment.startRow` and `.endRow` flat fields are removed; TypeScript callers must update to `files[i].sheets[0].startRow` etc. A new `FileSheetInfo` type is exported.

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
