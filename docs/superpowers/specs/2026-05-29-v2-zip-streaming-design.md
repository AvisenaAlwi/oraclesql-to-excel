# v2.0.0 Design: ZIP Streaming + Memory Safety

**Date:** 2026-05-29  
**Version target:** 2.0.0  
**Status:** Approved

---

## Background

### Current limitations

1. `.pipe()` and `.toBuffer()` throw if `.file()` is used — multi-file output only available via `.run()`.
2. `.run()` with `.file()` + `maxRowsPerFile` writes multiple `.xlsx` files to disk with no streaming path.
3. `showTotalRows` displays incorrect row range on files 2+ in multi-file exports (sheetIndex resets per file, globalOffset not tracked).
4. `backpressureThreshold` default (512 MB) is too high — RSS accumulates significantly before pausing Oracle fetch.

### Goal

1. Enable multi-file exports to be streamed as a single ZIP via `.pipe()`, written as a single `.zip` file via `.run()`, or returned as a ZIP Buffer via `.toBuffer()`.
2. Fix `showTotalRows` range calculation for multi-file exports.
3. Tighten default `backpressureThreshold` to reduce peak RSS.
4. Document memory behavior and mitigation clearly.

---

## Scenario clarification: 2.7M rows, maxRowsPerFile=1M, maxRowsPerSheet=500K

With a single `.sheet()` inside one `.file()`:

| File | Filename (in ZIP or on disk) | Sheet 1 | Sheet 2 |
|------|------------------------------|---------|---------|
| 1 | `report_1-1000000.xlsx` | rows 1–500K | rows 500K+1–1M |
| 2 | `report_1000001-2000000.xlsx` | rows 1M+1–1.5M | rows 1.5M+1–2M |
| 3 | `report_2000001-2700000.xlsx` | rows 2M+1–2.5M | rows 2.5M+1–2.7M |

= 3 files, 6 sheets. Oracle `ResultSet` opened once — carried across files via `openRS` + `overflowRows`.

---

## Architecture: Approach A — PassThrough → archiver streaming

```
OracleDB ResultSet
    ↓ getRows(fetchSize)              [only pause point]
ExcelJS WorkbookWriter
    ↓ XLSX chunks (streaming, zlib level 0)
PassThrough (one per file segment)
    ↓ appended as ZIP entry (dynamic, no pre-known size)
archiver npm (outer ZIP, store mode)
    ↓
[output target]
    ├─ .pipe(res)       → user's Writable stream
    ├─ .run()           → fs.WriteStream → .zip on disk
    └─ .toBuffer()      → PassThrough chunks → Buffer.concat()
```

Files are processed **sequentially**. archiver supports dynamic entry appending — no upfront declaration of file count required. Each PassThrough ends when `workbook.commit()` resolves, signalling archiver to close that ZIP entry and move on.

---

## API changes

### New builder method

```typescript
/**
 * Enable ZIP output mode. Required when using .file() with .pipe() or .toBuffer().
 * Optional for .run() — without it, .run() retains current behavior (multiple .xlsx files).
 *
 * When set, all terminal methods deliver output as a single ZIP archive:
 *   .pipe(res)    → streams ZIP to writable
 *   .run()        → writes <filePrefix>.zip to outputDir
 *   .toBuffer()   → returns ZIP as Buffer (avoid for large data — see memory notes)
 *
 * @param value - Default true when called without argument.
 */
asZip(value = true): this
```

### Terminal method behavior matrix

| Config | `.run()` | `.pipe()` | `.toBuffer()` |
|--------|----------|-----------|--------------|
| `.sheet()` only | single `.xlsx` — unchanged | single `.xlsx` stream — unchanged | `.xlsx` Buffer — unchanged |
| `.sheet()` only + `.asZip()` | `.asZip()` ignored — single `.xlsx` | `.asZip()` ignored — single `.xlsx` stream | `.asZip()` ignored — `.xlsx` Buffer |
| `.file()` + no `.asZip()` | multiple `.xlsx` → `MultiRunResult` — unchanged | **throw** (new informative message) | **throw** — unchanged |
| `.file()` + `.asZip()` | single `.zip` → `ZipRunResult` **(new)** | ZIP stream **(new)** | ZIP Buffer **(new)** |

### New type

```typescript
export interface ZipRunResult extends Result {
  /** Absolute path to the written .zip file. */
  file: string;
}
```

`.run()` return type becomes: `RunResult | MultiRunResult | ZipRunResult`.

### Error message update

When `.pipe()` or `.toBuffer()` is called with `.file()` but no `.asZip()`:

```
Error: .pipe() with .file() requires .asZip().
Call .asZip() on the builder so all files are streamed as a single ZIP.
Remember to set:
  Content-Type: application/zip
  Content-Disposition: attachment; filename="export.zip"
```

### Default value change

| Property | v1.x | v2.0 |
|----------|------|------|
| `backpressureThreshold` | 512 MB | **256 MB** |
| `fetchSize` | 50,000 | 50,000 (unchanged) |

---

## ZIP file naming

### Outer ZIP file (`.run()` only)

Uses `filePrefix` builder setting:

```js
.filePrefix('laporan').outputDir('/tmp').asZip().run()
// → /tmp/laporan.zip
```

### Entry names inside the ZIP

Determined by `FileConfig._name` passed to `.file()`:

| Scenario | Entry names inside ZIP |
|----------|----------------------|
| Single sheet, 1 segment | `{name}.xlsx` |
| Single sheet, split by maxRowsPerFile | `{name}_1-1000000.xlsx`, `{name}_1000001-2000000.xlsx` |
| Multi-sheet per file | `{name}_1.xlsx`, `{name}_2.xlsx` |
| Multiple `.file()` calls | `{name1}.xlsx`, `{name2}.xlsx` |

---

## Internal implementation

### New method: `_executeAsZip(outputTarget)`

Called from all three terminal methods when `_asZip === true && _files.length > 0`.

```typescript
async _executeAsZip(outputTarget: WorkbookTarget): Promise<Result> {
  const archive = archiver('zip', { store: true }); // store = no recompression
  // pipe to WriteStream (run), user Writable (pipe), or PassThrough (toBuffer)
  archive.pipe(outputTarget.stream!);

  for (const fileCfg of this._files) {
    let fileIndex = 0;
    // st = { openRS, pending, done, globalStart, globalEnd }
    // Loop continues while at least one sheet in this FileConfig is not done

    while ([...states.get(fileCfg)!.values()].some(st => !st.done)) {
      const pass = new PassThrough();
      archive.append(pass, { name: resolveZipEntryName(fileCfg, fileIndex, globalStart, globalEnd) });

      const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
        stream: pass,
        useStyles: true,
        useSharedStrings: false,
        zip: { zlib: { level: 0 } },
      });

      for (const sheetCfg of fileCfg._sheets) {
        const result = await this._executeSheetSegment(
          connection, workbook, sheetCfg,
          progressCtx, drainFn,
          maxRows, st.pending, st.openRS,
          st.globalStart - 1  // ← globalRowOffset (new param)
        );
        // update st
      }

      await workbook.commit(); // ends pass → archiver closes ZIP entry
      fileIndex++;
    }
  }

  await archive.finalize();
}
```

### `_executeSheetSegment` signature change

Add `globalRowOffset: number` as last parameter. Used only for `showTotalRows` range calculation:

```typescript
// Before (wrong for files 2+)
const start = sheetIndex * sheetCfg._maxRowsPerSheet + 1;
const end   = Math.min((sheetIndex + 1) * sheetCfg._maxRowsPerSheet, totalRows);

// After (correct)
const start = globalRowOffset + sheetIndex * sheetCfg._maxRowsPerSheet + 1;
const end   = Math.min(globalRowOffset + (sheetIndex + 1) * sheetCfg._maxRowsPerSheet, totalRows);
```

Existing callers (`_executeFileConfig`) pass `st.globalStart - 1`. Non-file path (`_execute`) passes `0`.

### drainFn for ZIP pipeline

For `.pipe()` + `.asZip()`: use **both** event-driven drain (write() backpressure on user's Writable) AND RSS fallback — identical to the existing `_execute` drainFn. The write() interception is applied on the user's Writable before piping it to archiver, so drain events propagate correctly.

For `.run()` + `.asZip()`: RSS-only polling (same as existing `.run()` since v1.1.4):

```typescript
const drainFn = async () => {
  if (process.memoryUsage().rss <= this._backpressureThreshold) return;
  const started = Date.now();
  while (process.memoryUsage().rss > this._backpressureThreshold) {
    if (Date.now() - started > 30_000) break;
    await new Promise<void>(r => setTimeout(r, 200));
  }
};
```

---

## Memory behavior and limitations

### How RSS polling protects memory

RSS polling is **reactive**, not preventive. It checks RSS after each batch completes, not during:

```
fetch batch (50K rows)
    ↓
ExcelJS processes batch → writes to archiver → output  [RSS rises here]
    ↓
drainFn() called → check RSS
    ↓
RSS > threshold? → PAUSE → wait for RSS to drop → fetch next batch
```

This prevents **unbounded accumulation** across batches — not instantaneous spike from a single batch.

### Memory comparison by terminal method

| Terminal | Data destination | OOM risk |
|----------|-----------------|----------|
| `.run()` + `.asZip()` | Disk (WriteStream) | Low — same as current `.run()` for XLSX |
| `.pipe()` + `.asZip()` | User Writable | Medium — RSS polling throttles slow-client scenarios |
| `.toBuffer()` + `.asZip()` | RAM | High for large data — entire ZIP in memory |

### `.toBuffer()` warning

`.toBuffer()` must hold the entire ZIP in memory before returning. For large exports (many rows, many files) this will OOM. Use `.run()` or `.pipe()` for large data.

### User mitigation for `.pipe()` + slow clients

When the end client downloads slowly and the output passes through a reverse proxy (nginx, etc.):
- Proxy accepts data from Node.js instantly (local socket) — write() always returns true
- RSS grows as archiver buffers outpace client download speed
- RSS polling pauses Oracle fetching once threshold is reached

To tune for your environment:

```js
OracleSqlToExcel()
  // Lower threshold if server has limited RAM
  .backpressureThreshold(128 * 1024 * 1024) // 128 MB

  // Lower fetchSize if rows are very wide (many columns, large text)
  .file('report', f => f
    .sheet('Data', s => s
      .fetchSize(10_000) // default 50K — reduce for wide rows
      .sql(SQL).columns(COLS)
    )
  )
  .asZip()
  .pipe(res)
```

---

## New dependency

```json
"dependencies": {
  "archiver": "^7.x"
},
"devDependencies": {
  "@types/archiver": "^6.x"
}
```

---

## Bug fix: `showTotalRows` incorrect range (multi-file)

**Root cause:** `sheetIndex` resets to 0 at the start of each `_executeSheetSegment` call. Without a global offset, the range calculation treats every file as if it starts at row 1.

**Symptom (screenshot confirmed):** File `Monitoring BMN_1000001-1500000.xlsx` displays "Showing rows 1 – 1,000,000 of 2,700,501 total" instead of the correct "Showing rows 1,000,001 – 1,500,000 of 2,700,501 total".

**Fix:** Pass `globalRowOffset = st.globalStart - 1` from `_executeFileConfig` into `_executeSheetSegment`. Both `_executeSheet` (non-file path, offset=0) and `_executeSheetSegment` (file path, offset=st.globalStart-1) receive this value.

---

## Backward compatibility

| Existing usage | v2.0 behavior |
|----------------|--------------|
| `.sheet()` + any terminal | Fully unchanged |
| `.file()` + `.run()` without `.asZip()` | Unchanged — multiple `.xlsx` files, `MultiRunResult` |
| `.file()` + `.pipe()` | Was: throw. Now: throw with better message (still throws) |
| `.file()` + `.toBuffer()` | Was: throw. Unchanged (still throws without `.asZip()`) |
| Manual `.backpressureThreshold(n)` | Unchanged — explicit value always wins |

---

## Breaking changes (v2.0.0)

1. **`backpressureThreshold` default 512 MB → 256 MB.** Users who rely on the default and have exports between 256–512 MB RSS will now pause more frequently. Users who set it manually are unaffected.
2. **`showTotalRows` output text changes.** Files 2+ now show correct row ranges instead of wrong ones. This is a fix, not a regression, but the displayed text changes.

---

## Versioning

**2.0.0** — scope justification:
- New output format (ZIP) and dependency (`archiver`)
- Default behavior change (`backpressureThreshold`)
- Two breaking changes noted above

---

## README additions

1. **ZIP streaming section** — examples for `.pipe()`, `.run()`, `.toBuffer()` with `.asZip()`
2. **Memory guide** — how RSS polling works, when OOM can still occur, mitigation (tune `fetchSize`, `backpressureThreshold`)
3. **`.toBuffer()` warning** — not recommended for large data
4. **Update method table** — add `asZip()` row, update `.run()` return type

---

## CHANGELOG v2.0.0 sections

```
### Added
- .asZip() builder method — enables ZIP output mode for .file() exports
- .pipe() + .file() + .asZip() — streams all files as a single ZIP to any Writable
- .run() + .file() + .asZip() — writes a single .zip file to outputDir
- .toBuffer() + .file() + .asZip() — returns ZIP as Buffer (not recommended for large data)
- ZipRunResult type exported for TypeScript callers

### Fixed
- showTotalRows incorrect row range on files 2+ in multi-file exports
  (sheetIndex reset per file without global offset — now passes globalRowOffset to _executeSheetSegment)
- .pipe() + .file() error message now explicitly mentions .asZip() and required HTTP headers

### Changed
- backpressureThreshold default: 512 MB → 256 MB (manual override unaffected)

### Breaking Changes
- backpressureThreshold default reduced from 512 MB to 256 MB
- showTotalRows display text corrected for multi-file exports (was showing wrong range)
```
