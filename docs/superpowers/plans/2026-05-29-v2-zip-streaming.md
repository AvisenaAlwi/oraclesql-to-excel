# v2.0.0 ZIP Streaming Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add `.asZip()` builder method that streams multi-file XLSX exports as a single ZIP through `.pipe()`, `.run()`, and `.toBuffer()`; fix `showTotalRows` range bug in multi-file exports; tighten `backpressureThreshold` default.

**Architecture:** When `.asZip()` is set and `.file()` is used, a new `_executeAsZip()` method drives all three terminal methods. It creates an `archiver` ZIP stream and, for each file segment, pipes an ExcelJS WorkbookWriter through a PassThrough into a named ZIP entry — all streamed sequentially without temp files on disk. The existing `_execute()` and `_executeFileConfig()` paths are untouched for non-ZIP usage.

**Tech Stack:** TypeScript, ExcelJS (existing), archiver ^5.3.2 (already a transitive dep — declare as direct dep), adm-zip (new devDep for test ZIP reading), Jest + ts-jest.

**Spec:** `docs/superpowers/specs/2026-05-29-v2-zip-streaming-design.md`

---

## File map

| File | Change |
|------|--------|
| `package.json` | Add `archiver` to dependencies, `adm-zip` + `@types/adm-zip` to devDependencies |
| `src/index.ts` | All source changes (types, builder field, new method, terminal method updates, default change) |
| `tests/helpers.ts` | Add `readZipBuffer` + `readRawWorksheet` helpers |
| `tests/zip.test.ts` | New file — all ZIP-specific tests |
| `tests/builder.test.ts` | Add asZip() chainability test |
| `README.md` | Add ZIP streaming section + memory guide |
| `CHANGELOG.md` | Add v2.0.0 entry |

---

## Task 1: Install dependencies

**Files:**
- Modify: `package.json`

- [ ] **Step 1: Declare archiver as direct dependency and add adm-zip devDep**

In `package.json`, update `dependencies` and `devDependencies`:

```json
"dependencies": {
  "archiver": "^5.3.2",
  "exceljs": "^4.4.0"
},
"devDependencies": {
  "@types/adm-zip": "^0.5.5",
  "@types/archiver": "^5.3.4",
  "@types/jest": "^30.0.0",
  "@types/node": "^25.9.1",
  "adm-zip": "^0.5.10",
  "jest": "^30.4.2",
  "ts-jest": "^29.4.11",
  "typescript": "^5.0.0"
}
```

- [ ] **Step 2: Install**

```bash
npm install
```

Expected: no errors. `archiver`, `adm-zip`, `@types/archiver`, `@types/adm-zip` visible in `node_modules/`.

- [ ] **Step 3: Verify archiver import works**

```bash
node -e "const a = require('archiver'); const arc = a('zip', { zlib: { level: 0 } }); console.log('OK', typeof arc.append)"
```

Expected: `OK function`

- [ ] **Step 4: Commit**

```bash
git add package.json package-lock.json
git commit -m "chore: add archiver as direct dep, adm-zip as devDep for v2.0.0"
```

---

## Task 2: Scaffold — types, builder field, helpers

**Files:**
- Modify: `src/index.ts`
- Modify: `tests/helpers.ts`

- [ ] **Step 1: Add `ZipRunResult` interface after `BufferResult` (~line 137)**

In `src/index.ts`, after the `BufferResult` interface:

```typescript
export interface ZipRunResult extends Result {
  /** Absolute path to the written .zip file. */
  file: string;
}
```

- [ ] **Step 2: Add `import archiver from 'archiver'` at top of src/index.ts**

After the existing imports:

```typescript
import fs from 'fs';
import ExcelJS from 'exceljs';
import path from 'path';
import { PassThrough, Writable } from 'stream';
import archiver from 'archiver';
```

- [ ] **Step 3: Add `_asZip` field to `OracleSqlToExcelBuilder` constructor**

In the class fields (after `_backpressureThreshold`):

```typescript
/** @private */ _asZip : boolean;
```

In `constructor()`, after `this._backpressureThreshold = 512 * 1024 * 1024;`:

```typescript
this._asZip = false;
```

- [ ] **Step 4: Add `asZip()` builder method**

After the `backpressureThreshold()` method:

```typescript
/**
 * Enable ZIP output mode for `.file()` exports.
 * Required when calling `.pipe()` or `.toBuffer()` with `.file()`.
 * Optional for `.run()` — without it, `.run()` retains existing behavior (multiple `.xlsx` files).
 *
 * When set, all terminal methods deliver a single ZIP archive:
 *   `.pipe(res)`    → streams ZIP to writable
 *   `.run()`        → writes `<filePrefix>.zip` to `outputDir`
 *   `.toBuffer()`   → returns ZIP as Buffer (avoid for large data)
 *
 * Has no effect when `.file()` is not used.
 * @param value - Default `true` when called without argument.
 */
asZip(value = true): this { this._asZip = value; return this; }
```

- [ ] **Step 5: Add `readZipBuffer` and `readRawWorksheet` to `tests/helpers.ts`**

```typescript
import AdmZip from 'adm-zip';

/**
 * Extract all entries from a ZIP buffer. Returns name + raw data per entry.
 */
export function readZipBuffer(
  zipBuffer: Buffer
): Array<{ name: string; data: Buffer }> {
  const zip = new AdmZip(zipBuffer);
  return zip.getEntries().map((e) => ({
    name: e.entryName,
    data: e.getData(),
  }));
}

/**
 * Read ALL rows (including pre-header rows like docHeader / showTotalRows) from a
 * named worksheet in an XLSX buffer. Returns 1-indexed ExcelJS cell value arrays.
 */
export async function readRawWorksheet(
  buffer : Buffer,
  sheetName: string
): Promise<Array<Array<unknown>>> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer as unknown as Buffer);
  const ws = wb.getWorksheet(sheetName);
  if (!ws) return [];
  const rows: Array<Array<unknown>> = [];
  ws.eachRow((row) => { rows.push(row.values as Array<unknown>); });
  return rows;
}
```

- [ ] **Step 6: Verify TypeScript compiles with no errors**

```bash
npx tsc --noEmit
```

Expected: no errors.

- [ ] **Step 7: Commit**

```bash
git add src/index.ts tests/helpers.ts
git commit -m "feat: scaffold ZipRunResult type, asZip() builder method, zip test helpers"
```

---

## Task 3: Fix `showTotalRows` globalRowOffset bug (TDD)

**Files:**
- Modify: `src/index.ts` (`_executeSheetSegment` signature + body, `_executeFileConfig` call site)
- Modify: `tests/zip.test.ts` (create file)

- [ ] **Step 1: Create `tests/zip.test.ts` with the failing globalRowOffset test**

```typescript
import fs from 'fs';
import os from 'os';
import path from 'path';
import ExcelJS from 'exceljs';
import { OracleSqlToExcel } from '../src/index';
import { createStreamConn, createCountConn, readZipBuffer, readBuffer, readRawWorksheet } from './helpers';

const COLS = [
  { key: 'ID',   header: 'ID',   type: 'number' as const },
  { key: 'NAME', header: 'Name', type: 'text'   as const },
];

function makeRows(n: number): Record<string, unknown>[] {
  return Array.from({ length: n }, (_, i) => ({ ID: i + 1, NAME: `Row${i + 1}` }));
}

// Factory that returns countConn for COUNT queries, streamConn for data queries.
// The library calls connectionFactory: once for the main stream + once per showTotalRows sheet.
// Count query always comes AFTER the main connection is opened in _executeFileConfig.
function makeCountThenStreamFactory(rows: Record<string, unknown>[], total: number) {
  const conns = [
    createStreamConn(rows, COLS.map((c) => ({ name: c.key }))), // main connection
    createCountConn(total),                                       // count connection
  ];
  let i = 0;
  return () => Promise.resolve(conns[i++] ?? createStreamConn(rows, COLS.map((c) => ({ name: c.key }))));
}

// ── globalRowOffset fix ───────────────────────────────────────────────────────

describe('showTotalRows — globalRowOffset fix (multi-file)', () => {
  it('file 2 shows correct row range, not 1-N', async () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ote-test-'));
    const rows   = makeRows(5); // 5 total rows
    // maxRowsPerFile=3 → file 1: rows 1-3, file 2: rows 4-5
    const factory = makeCountThenStreamFactory(rows, 5);

    const result = await OracleSqlToExcel()
      .connectionFactory(factory)
      .outputDir(tmpDir)
      .file('data', (f) => f
        .maxRowsPerFile(3)
        .sheet('Sheet1', (s) => s
          .sql('SELECT * FROM T')
          .columns(COLS)
          .showTotalRows()
        )
      )
      .run() as any;

    expect(result.success).toBe(true);
    expect(result.files).toHaveLength(2);

    // Read file 2 (rows 4-5)
    const file2Buffer = fs.readFileSync(result.files[1].file);
    const rawRows     = await readRawWorksheet(file2Buffer, 'Sheet1');
    // rawRows[0] is undefined (1-indexed), rawRows[1] is first row = showTotalRows line
    const showTotalRowsText = String(rawRows[1]?.[1] ?? '');
    // Must show "4" as start, not "1"
    expect(showTotalRowsText).toMatch(/4/);
    expect(showTotalRowsText).toMatch(/5/);
    expect(showTotalRowsText).toMatch(/5 total/i);
    expect(showTotalRowsText).not.toMatch(/^Showing rows 1/);

    fs.rmSync(tmpDir, { recursive: true });
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

```bash
npx jest tests/zip.test.ts --testNamePattern="globalRowOffset" --no-coverage
```

Expected: FAIL — the text shows "Showing rows 1" instead of "Showing rows 4".

- [ ] **Step 3: Add `globalRowOffset` parameter to `_executeSheetSegment`**

In `src/index.ts`, find the `_executeSheetSegment` method signature (~line 1063) and add the parameter with a default:

```typescript
async _executeSheetSegment(
  connection      : OracleConnection,
  workbook        : StreamWorkbook,
  sheetCfg        : SheetConfig,
  progressCtx     : ProgressCtx,
  drainFn         : (() => Promise<void>) | null,
  maxRows         : number,
  pendingRows     : Record<string, unknown>[],
  existingRS      : OracleResultSet | null,
  globalRowOffset : number = 0
): Promise<SheetSegmentResult> {
```

- [ ] **Step 4: Fix the showTotalRows range calculation in `_executeSheetSegment`**

Find this block inside `_executeSheetSegment`'s `createNewSheet` (~line 1104):

```typescript
if (sheetCfg._resolvedTotalRows != null) {
  const fmt      = (n: number): string => n.toLocaleString('en-US');
  const start    = sheetIndex * sheetCfg._maxRowsPerSheet + 1;
  const end      = Math.min((sheetIndex + 1) * sheetCfg._maxRowsPerSheet, sheetCfg._resolvedTotalRows);
```

Replace with:

```typescript
if (sheetCfg._resolvedTotalRows != null) {
  const fmt      = (n: number): string => n.toLocaleString('en-US');
  const start    = globalRowOffset + sheetIndex * sheetCfg._maxRowsPerSheet + 1;
  const end      = Math.min(globalRowOffset + (sheetIndex + 1) * sheetCfg._maxRowsPerSheet, sheetCfg._resolvedTotalRows);
```

- [ ] **Step 5: Pass `globalRowOffset` from `_executeFileConfig` call site**

Find the `_executeSheetSegment` call inside `_executeFileConfig` (~line 1297):

```typescript
const seg = await this._executeSheetSegment(
  connection, workbook, sheetCfg, progressCtx, drainFn,
  maxRows, st.pending, st.openRS
);
```

Replace with:

```typescript
const seg = await this._executeSheetSegment(
  connection, workbook, sheetCfg, progressCtx, drainFn,
  maxRows, st.pending, st.openRS,
  st.globalStart - 1
);
```

- [ ] **Step 6: Run the test to verify it passes**

```bash
npx jest tests/zip.test.ts --testNamePattern="globalRowOffset" --no-coverage
```

Expected: PASS

- [ ] **Step 7: Run all existing tests to confirm no regressions**

```bash
npx jest --no-coverage
```

Expected: all tests pass.

- [ ] **Step 8: Commit**

```bash
git add src/index.ts tests/zip.test.ts
git commit -m "fix: correct showTotalRows row range on files 2+ in multi-file exports (globalRowOffset)"
```

---

## Task 4: TDD — `_executeAsZip` via `.toBuffer()` + `.asZip()`

**Files:**
- Modify: `tests/zip.test.ts`
- Modify: `src/index.ts`

- [ ] **Step 1: Add failing tests for `.toBuffer()` + `.asZip()`**

Append to `tests/zip.test.ts`:

```typescript
// ── toBuffer + asZip ──────────────────────────────────────────────────────────

describe('toBuffer() + asZip()', () => {
  const rows = makeRows(5);
  const factory = () => Promise.resolve(createStreamConn(rows, COLS.map((c) => ({ name: c.key }))));

  it('returns success=true and non-empty buffer', async () => {
    const { success, buffer } = await OracleSqlToExcel()
      .connectionFactory(factory)
      .file('report', (f) => f
        .maxRowsPerFile(3)
        .sheet('Data', (s) => s.sql('SELECT * FROM T').columns(COLS))
      )
      .asZip()
      .toBuffer();

    expect(success).toBe(true);
    expect(buffer.length).toBeGreaterThan(0);
  });

  it('ZIP buffer contains correct number of entries', async () => {
    const { buffer } = await OracleSqlToExcel()
      .connectionFactory(factory)
      .file('report', (f) => f
        .maxRowsPerFile(3)
        .sheet('Data', (s) => s.sql('SELECT * FROM T').columns(COLS))
      )
      .asZip()
      .toBuffer();

    const entries = readZipBuffer(buffer);
    // 5 rows, maxRowsPerFile=3 → 2 entries
    expect(entries).toHaveLength(2);
  });

  it('ZIP entry names are sequential', async () => {
    const { buffer } = await OracleSqlToExcel()
      .connectionFactory(factory)
      .file('myreport', (f) => f
        .maxRowsPerFile(3)
        .sheet('Data', (s) => s.sql('SELECT * FROM T').columns(COLS))
      )
      .asZip()
      .toBuffer();

    const entries = readZipBuffer(buffer);
    expect(entries[0].name).toBe('myreport_1.xlsx');
    expect(entries[1].name).toBe('myreport_2.xlsx');
  });

  it('each XLSX entry contains correct row data', async () => {
    const { buffer } = await OracleSqlToExcel()
      .connectionFactory(factory)
      .file('rep', (f) => f
        .maxRowsPerFile(3)
        .sheet('Data', (s) => s.sql('SELECT * FROM T').columns(COLS))
      )
      .asZip()
      .toBuffer();

    const entries = readZipBuffer(buffer);

    const [sheet1] = await readBuffer(entries[0].data, 'Data');
    expect(sheet1.rows).toHaveLength(3); // rows 1-3

    const [sheet2] = await readBuffer(entries[1].data, 'Data');
    expect(sheet2.rows).toHaveLength(2); // rows 4-5
  });

  it('returns success=false when connectionFactory fails', async () => {
    const { success, buffer, error } = await OracleSqlToExcel()
      .connectionFactory(() => Promise.reject(new Error('DB down')))
      .file('rep', (f) => f.sheet('Data', (s) => s.sql('SELECT 1').columns(COLS)))
      .asZip()
      .toBuffer();

    expect(success).toBe(false);
    expect(buffer.length).toBe(0);
    expect(error).toMatch(/DB down/i);
  });

  it('throws when .file() used without .asZip()', async () => {
    await expect(
      OracleSqlToExcel()
        .connectionFactory(factory)
        .file('rep', (f) => f.sheet('Data', (s) => s.sql('SELECT 1').columns(COLS)))
        .toBuffer()
    ).rejects.toThrow(/asZip/i);
  });

  it('single file (data fits in maxRowsPerFile) produces 1 ZIP entry', async () => {
    const { buffer } = await OracleSqlToExcel()
      .connectionFactory(factory)
      .file('single', (f) => f
        .maxRowsPerFile(100)
        .sheet('Data', (s) => s.sql('SELECT * FROM T').columns(COLS))
      )
      .asZip()
      .toBuffer();

    const entries = readZipBuffer(buffer);
    expect(entries).toHaveLength(1);
    expect(entries[0].name).toBe('single_1.xlsx');
  });
});
```

- [ ] **Step 2: Run to verify all new tests fail**

```bash
npx jest tests/zip.test.ts --testNamePattern="toBuffer.*asZip|asZip.*toBuffer" --no-coverage 2>&1 | tail -20
```

Expected: FAIL — "not a function" or similar.

- [ ] **Step 3: Implement `_executeAsZip` method in `src/index.ts`**

Add this method to `OracleSqlToExcelBuilder` after `_executeFileConfig`:

```typescript
/**
 * Core execution for .asZip() mode. Streams each FileConfig's XLSX output
 * sequentially into named ZIP entries via archiver.
 * @private
 */
async _executeAsZip(
  outputStream : Writable,
  drainFn      : (() => Promise<void>) | null
): Promise<Result> {
  const allSheets   : string[]    = [];
  const progressCtx : ProgressCtx = { totalRowsWritten: 0 };
  let   totalSkipped               = 0;

  const isDevEnv = !['production', 'prod'].includes((process.env.NODE_ENV ?? '').toLowerCase());
  const dbg      = (msg: string): void => { if (this._debug && isDevEnv) console.log(`[OracleSqlToExcel:DEBUG] ${msg}`); };

  const archive = archiver('zip', { zlib: { level: 0 } });
  archive.pipe(outputStream);

  let archiveError: Error | null = null;
  archive.on('error', (err: Error) => { archiveError = err; });

  try {
    if (!this._connectionFactory) {
      throw new Error('No connection factory set. Call .connectionFactory(() => pool.getConnection()) before running the export.');
    }

    for (const fileCfg of this._files) {
      let connection: OracleConnection | null = null;

      try {
        connection = await this._connectionFactory();

        // COUNT queries for showTotalRows — parallel, each on its own connection
        await Promise.all(fileCfg._sheets.map(async (sheetCfg) => {
          const label = Array.isArray(sheetCfg._name) ? sheetCfg._name[0] : sheetCfg._name;
          if (!sheetCfg._showTotalRows) return;
          try {
            const trimmed = sheetCfg._sql.trim().replace(/;+$/, '');
            if (/^\s*WITH\s+/i.test(trimmed)) return;
            const countConn = await this._connectionFactory!();
            try {
              const result = await countConn.execute(
                `SELECT COUNT(*) AS TOTAL FROM (${trimmed})`,
                sheetCfg._param,
                { outFormat: OUT_FORMAT_OBJECT }
              );
              sheetCfg._resolvedTotalRows = (result.rows?.[0]?.['TOTAL'] as number | undefined) ?? null;
            } finally {
              await countConn.close().catch(() => {});
            }
          } catch (err) {
            const msg = err instanceof Error ? err.message : String(err);
            console.error(`[OracleSqlToExcel:ERROR] COUNT query failed for sheet "${label}": ${msg}`);
          }
        }));

        const maxRows = fileCfg._maxRowsPerFile > 0 ? fileCfg._maxRowsPerFile : Number.MAX_SAFE_INTEGER;

        interface SheetState {
          openRS     : OracleResultSet | null;
          pending    : Record<string, unknown>[];
          done       : boolean;
          globalStart: number;
          globalEnd  : number;
        }

        const states = new Map<SheetConfig, SheetState>(
          fileCfg._sheets.map((s) => [s, { openRS: null, pending: [], done: false, globalStart: 1, globalEnd: 0 }])
        );

        let fileIndex = 0;

        while ([...states.values()].some((st) => !st.done)) {
          if (archiveError) throw archiveError;

          const entryName = `${fileCfg._name}_${fileIndex + 1}.xlsx`;
          dbg(`ZIP: append entry "${entryName}"`);

          const pass = new PassThrough();
          archive.append(pass, { name: entryName });

          const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
            stream          : pass,
            useStyles       : true,
            useSharedStrings: false,
            zip             : { zlib: { level: 0 } },
          } as unknown as ExcelJS.stream.xlsx.WorkbookWriterOptions);

          for (const sheetCfg of fileCfg._sheets) {
            const st = states.get(sheetCfg)!;
            if (st.done) continue;

            const seg = await this._executeSheetSegment(
              connection, workbook, sheetCfg, progressCtx, drainFn,
              maxRows, st.pending, st.openRS,
              st.globalStart - 1
            );

            st.globalEnd   = st.globalStart + seg.rowsWritten - 1;
            allSheets.push(...seg.sheetNames);
            totalSkipped  += seg.skippedRows;
            st.globalStart = st.globalEnd + 1;
            st.openRS      = seg.openRS;
            st.pending     = seg.overflowRows;
            st.done        = seg.openRS === null && seg.overflowRows.length === 0;
          }

          await workbook.commit(); // ends pass → archiver closes this ZIP entry
          dbg(`ZIP: entry "${entryName}" committed`);
          fileIndex++;
        }
      } finally {
        if (connection) await connection.close().catch(() => {});
      }
    }

    if (archiveError) throw archiveError;

    dbg('ZIP: archive.finalize()');
    const finalizePromise = new Promise<void>((resolve, reject) => {
      archive.on('finish', resolve);
      archive.on('error', reject);
    });
    archive.finalize();
    await finalizePromise;
    dbg('ZIP: archive done');

    return { success: true, sheets: allSheets, skippedRows: totalSkipped };
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error(`[OracleSqlToExcel:ERROR] _executeAsZip FAILED — ${msg}`);
    if (!outputStream.writableEnded) {
      try { outputStream.end(); } catch (_) {}
    }
    return { success: false, error: msg, sheets: allSheets, skippedRows: totalSkipped };
  }
}
```

- [ ] **Step 4: Update `.toBuffer()` to route through `_executeAsZip`**

In the existing `toBuffer()` method, replace:

```typescript
if (this._files.length > 0) {
  throw new Error('.toBuffer() does not support .file() — multi-file output requires .run().');
}
```

With:

```typescript
if (this._files.length > 0) {
  if (!this._asZip) {
    throw new Error(
      '.toBuffer() with .file() requires .asZip().\n' +
      'Call .asZip() on the builder so all files are returned as a single ZIP Buffer.\n' +
      'For large data, prefer .run() or .pipe() to avoid loading the entire ZIP in memory.'
    );
  }
  const zipChunks: Buffer[] = [];
  const zipPass             = new PassThrough();
  zipPass.on('data', (chunk: Buffer) => zipChunks.push(chunk));
  const zipDone = new Promise<void>((resolve, reject) => {
    zipPass.on('finish', resolve);
    zipPass.on('error', reject);
  });
  const result = await this._executeAsZip(zipPass, null);
  if (!zipPass.writableEnded) zipPass.end();
  await zipDone.catch(() => {});
  return { ...result, buffer: result.success ? Buffer.concat(zipChunks) : Buffer.alloc(0) };
}
```

- [ ] **Step 5: Run the new tests to verify they pass**

```bash
npx jest tests/zip.test.ts --testNamePattern="toBuffer.*asZip|asZip.*toBuffer" --no-coverage
```

Expected: all 7 tests PASS.

- [ ] **Step 6: Run all tests**

```bash
npx jest --no-coverage
```

Expected: all pass.

- [ ] **Step 7: Commit**

```bash
git add src/index.ts tests/zip.test.ts
git commit -m "feat: implement _executeAsZip + toBuffer() ZIP path"
```

---

## Task 5: TDD — `.pipe()` + `.asZip()`

**Files:**
- Modify: `tests/zip.test.ts`
- Modify: `src/index.ts`

- [ ] **Step 1: Add failing tests**

Append to `tests/zip.test.ts`:

```typescript
// ── pipe + asZip ──────────────────────────────────────────────────────────────

describe('pipe() + asZip()', () => {
  const rows    = makeRows(5);
  const factory = () => Promise.resolve(createStreamConn(rows, COLS.map((c) => ({ name: c.key }))));

  it('streams a valid ZIP to a PassThrough', async () => {
    const chunks : Buffer[] = [];
    const pass              = new PassThrough();
    pass.on('data', (c: Buffer) => chunks.push(c));
    const done = new Promise<void>((res, rej) => { pass.on('finish', res); pass.on('error', rej); });

    const result = await OracleSqlToExcel()
      .connectionFactory(factory)
      .file('rep', (f) => f
        .maxRowsPerFile(3)
        .sheet('Data', (s) => s.sql('SELECT * FROM T').columns(COLS))
      )
      .asZip()
      .pipe(pass);

    await done;
    expect(result.success).toBe(true);

    const zipBuffer = Buffer.concat(chunks);
    const entries   = readZipBuffer(zipBuffer);
    expect(entries).toHaveLength(2);
    expect(entries[0].name).toBe('rep_1.xlsx');
    expect(entries[1].name).toBe('rep_2.xlsx');
  });

  it('throws when .file() used without .asZip()', async () => {
    const pass = new PassThrough();
    await expect(
      OracleSqlToExcel()
        .connectionFactory(factory)
        .file('rep', (f) => f.sheet('Data', (s) => s.sql('SELECT 1').columns(COLS)))
        .pipe(pass)
    ).rejects.toThrow(/asZip/i);
  });

  it('error message mentions Content-Type hint', async () => {
    const pass = new PassThrough();
    try {
      await OracleSqlToExcel()
        .connectionFactory(factory)
        .file('rep', (f) => f.sheet('Data', (s) => s.sql('SELECT 1').columns(COLS)))
        .pipe(pass);
    } catch (e) {
      expect(String(e)).toMatch(/Content-Type/i);
    }
  });
});
```

- [ ] **Step 2: Run to verify tests fail**

```bash
npx jest tests/zip.test.ts --testNamePattern="pipe.*asZip|asZip.*pipe" --no-coverage 2>&1 | tail -15
```

Expected: FAIL.

- [ ] **Step 3: Update `.pipe()` in `src/index.ts`**

Replace the existing `.pipe()` method body:

```typescript
async pipe(writableStream: Writable): Promise<Result> {
  if (this._files.length > 0 && !this._asZip) {
    throw new Error(
      '.pipe() with .file() requires .asZip().\n' +
      'Call .asZip() on the builder so all files are streamed as a single ZIP.\n' +
      'Remember to set Content-Type: application/zip and Content-Disposition: attachment; filename="export.zip" before piping.'
    );
  }
  if (this._files.length > 0 && this._asZip) {
    const rssThreshold = this._backpressureThreshold;
    const drainFn: (() => Promise<void>) | null = rssThreshold > 0
      ? async () => {
          if (process.memoryUsage().rss <= rssThreshold) return;
          const started = Date.now();
          while (process.memoryUsage().rss > rssThreshold) {
            if (Date.now() - started > 30_000) break;
            await new Promise<void>((r) => setTimeout(r, 200));
          }
        }
      : null;
    return this._executeAsZip(writableStream, drainFn);
  }
  return this._execute({ stream: writableStream });
}
```

- [ ] **Step 4: Run the pipe+asZip tests**

```bash
npx jest tests/zip.test.ts --testNamePattern="pipe.*asZip|asZip.*pipe" --no-coverage
```

Expected: all PASS.

- [ ] **Step 5: Run all tests**

```bash
npx jest --no-coverage
```

Expected: all pass.

- [ ] **Step 6: Commit**

```bash
git add src/index.ts tests/zip.test.ts
git commit -m "feat: pipe() ZIP path + updated error message for pipe() + file() without asZip()"
```

---

## Task 6: TDD — `.run()` + `.asZip()`

**Files:**
- Modify: `tests/zip.test.ts`
- Modify: `src/index.ts`

- [ ] **Step 1: Add failing tests**

Append to `tests/zip.test.ts`:

```typescript
// ── run + asZip ───────────────────────────────────────────────────────────────

describe('run() + asZip()', () => {
  const rows    = makeRows(5);
  const factory = () => Promise.resolve(createStreamConn(rows, COLS.map((c) => ({ name: c.key }))));

  it('writes a .zip file to outputDir and returns ZipRunResult', async () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ote-zip-'));
    const result: any = await OracleSqlToExcel()
      .connectionFactory(factory)
      .outputDir(tmpDir)
      .filePrefix('export')
      .file('rep', (f) => f
        .maxRowsPerFile(3)
        .sheet('Data', (s) => s.sql('SELECT * FROM T').columns(COLS))
      )
      .asZip()
      .run();

    expect(result.success).toBe(true);
    expect(result.file).toBe(path.join(tmpDir, 'export.zip'));
    expect(fs.existsSync(result.file)).toBe(true);

    const zipBuffer = fs.readFileSync(result.file);
    const entries   = readZipBuffer(zipBuffer);
    expect(entries).toHaveLength(2);
    expect(entries[0].name).toBe('rep_1.xlsx');

    fs.rmSync(tmpDir, { recursive: true });
  });

  it('ZIP contains correct XLSX data', async () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ote-zip-'));
    const result: any = await OracleSqlToExcel()
      .connectionFactory(factory)
      .outputDir(tmpDir)
      .filePrefix('exp')
      .file('data', (f) => f
        .maxRowsPerFile(3)
        .sheet('Sheet1', (s) => s.sql('SELECT * FROM T').columns(COLS))
      )
      .asZip()
      .run();

    const zipBuffer = fs.readFileSync(result.file);
    const entries   = readZipBuffer(zipBuffer);

    const [s1] = await readBuffer(entries[0].data, 'Sheet1');
    expect(s1.rows).toHaveLength(3);
    expect(s1.rows[0][0]).toBe(1);

    const [s2] = await readBuffer(entries[1].data, 'Sheet1');
    expect(s2.rows).toHaveLength(2);
    expect(s2.rows[0][0]).toBe(4);

    fs.rmSync(tmpDir, { recursive: true });
  });

  it('deletes partial zip on error', async () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ote-zip-'));
    const result: any = await OracleSqlToExcel()
      .connectionFactory(() => Promise.reject(new Error('DB fail')))
      .outputDir(tmpDir)
      .filePrefix('bad')
      .file('rep', (f) => f.sheet('Data', (s) => s.sql('SELECT 1').columns(COLS)))
      .asZip()
      .run();

    expect(result.success).toBe(false);
    // zip file should not exist (cleaned up)
    expect(fs.existsSync(path.join(tmpDir, 'bad.zip'))).toBe(false);

    fs.rmSync(tmpDir, { recursive: true });
  });

  it('.run() without .asZip() still produces multiple .xlsx files (backward compat)', async () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ote-zip-'));
    const result: any = await OracleSqlToExcel()
      .connectionFactory(factory)
      .outputDir(tmpDir)
      .file('data', (f) => f
        .maxRowsPerFile(3)
        .sheet('Sheet1', (s) => s.sql('SELECT * FROM T').columns(COLS))
      )
      .run();

    expect(result.success).toBe(true);
    expect(result.files).toHaveLength(2); // MultiRunResult, not ZipRunResult
    expect(result.files[0].file).toMatch(/\.xlsx$/);

    fs.rmSync(tmpDir, { recursive: true });
  });
});
```

- [ ] **Step 2: Run to verify tests fail**

```bash
npx jest tests/zip.test.ts --testNamePattern="run.*asZip|asZip.*run" --no-coverage 2>&1 | tail -15
```

Expected: FAIL.

- [ ] **Step 3: Update `.run()` in `src/index.ts`**

Change the `.run()` return type and add the asZip path before the existing multi-file block:

```typescript
async run(): Promise<RunResult | MultiRunResult | ZipRunResult> {
  // ── ZIP path ────────────────────────────────────────────────────────────────
  if (this._files.length > 0 && this._asZip) {
    const zipFile = path.join(this._outputDir, `${this._filePrefix}.zip`);
    const ws      = fs.createWriteStream(zipFile);
    const wsDone  = new Promise<void>((resolve, reject) => {
      ws.on('finish', resolve);
      ws.on('error',  reject);
    });
    const rssThreshold = this._backpressureThreshold;
    const drainFn: (() => Promise<void>) | null = rssThreshold > 0
      ? async () => {
          if (process.memoryUsage().rss <= rssThreshold) return;
          const started = Date.now();
          while (process.memoryUsage().rss > rssThreshold) {
            if (Date.now() - started > 30_000) break;
            await new Promise<void>((r) => setTimeout(r, 200));
          }
        }
      : null;
    const result = await this._executeAsZip(ws, drainFn);
    await wsDone.catch(() => {});
    if (!result.success) {
      fs.promises.unlink(zipFile).catch(() => {});
    }
    return { ...result, file: zipFile } as ZipRunResult;
  }

  // ── Existing multi-file path (.file() without .asZip()) ─────────────────────
  if (this._files.length > 0) {
    // ... existing code unchanged ...
  }

  // ── Existing single-file path ────────────────────────────────────────────────
  // ... existing code unchanged ...
}
```

- [ ] **Step 4: Run the run+asZip tests**

```bash
npx jest tests/zip.test.ts --testNamePattern="run.*asZip|asZip.*run" --no-coverage
```

Expected: all PASS.

- [ ] **Step 5: Run all tests**

```bash
npx jest --no-coverage
```

Expected: all pass.

- [ ] **Step 6: Commit**

```bash
git add src/index.ts tests/zip.test.ts
git commit -m "feat: run() ZIP path — writes single .zip to outputDir when asZip() is set"
```

---

## Task 7: TDD — builder validation + chainability

**Files:**
- Modify: `tests/builder.test.ts`
- Modify: `tests/zip.test.ts`

- [ ] **Step 1: Add `asZip()` to the chainability test in `tests/builder.test.ts`**

Find the `'builder methods are chainable'` test and add `.asZip()`:

```typescript
it('builder methods are chainable', () => {
  const builder = OracleSqlToExcel()
    .connectionFactory(conn)
    .outputDir('/tmp')
    .filePrefix('test')
    .compress(true)
    .debug(false)
    .asZip(true)
    .onProgress(() => {})
    .executeOptions({ autoCommit: true })
    .sheet('S', (s) => s.sql('SELECT 1 FROM DUAL').freezeHeader().autoFilter());
  expect(builder).toBeDefined();
});
```

- [ ] **Step 2: Add asZip ignored without .file() test**

Append to `tests/zip.test.ts`:

```typescript
// ── asZip() ignored without .file() ──────────────────────────────────────────

describe('asZip() without .file()', () => {
  const rows    = makeRows(2);
  const factory = () => Promise.resolve(createStreamConn(rows, COLS.map((c) => ({ name: c.key }))));

  it('toBuffer() returns plain XLSX (not ZIP) when asZip() used with .sheet() only', async () => {
    const { success, buffer } = await OracleSqlToExcel()
      .connectionFactory(factory)
      .sheet('Data', (s) => s.sql('SELECT * FROM T').columns(COLS))
      .asZip() // ignored — no .file()
      .toBuffer();

    expect(success).toBe(true);
    // A ZIP starts with PK (0x50 0x4B). XLSX is also a ZIP internally,
    // but ExcelJS returns a valid XLSX. Verify it can be parsed as XLSX:
    const [{ rows: parsed }] = await readBuffer(buffer);
    expect(parsed).toHaveLength(2);
  });
});
```

- [ ] **Step 3: Run to verify tests pass (asZip is already a no-op for non-file path)**

```bash
npx jest tests/builder.test.ts tests/zip.test.ts --testNamePattern="chainable|asZip.*without" --no-coverage
```

Expected: PASS (asZip() already returns `this` and the non-file code paths ignore `_asZip`).

- [ ] **Step 4: Run all tests**

```bash
npx jest --no-coverage
```

Expected: all pass.

- [ ] **Step 5: Commit**

```bash
git add tests/builder.test.ts tests/zip.test.ts
git commit -m "test: asZip() chainability and no-op behavior without .file()"
```

---

## Task 8: Change `backpressureThreshold` default

**Files:**
- Modify: `src/index.ts`
- Modify: `tests/builder.test.ts`

- [ ] **Step 1: Add a failing test for the new default**

Append to `tests/builder.test.ts`:

```typescript
it('backpressureThreshold default is 256 MB', () => {
  const builder = OracleSqlToExcel() as any;
  expect(builder._backpressureThreshold).toBe(256 * 1024 * 1024);
});
```

- [ ] **Step 2: Run to verify it fails**

```bash
npx jest tests/builder.test.ts --testNamePattern="backpressureThreshold" --no-coverage
```

Expected: FAIL — current value is 512 MB.

- [ ] **Step 3: Change the default in `src/index.ts`**

In `OracleSqlToExcelBuilder` constructor, find:

```typescript
this._backpressureThreshold = 512 * 1024 * 1024; // 512 MB
```

Replace with:

```typescript
this._backpressureThreshold = 256 * 1024 * 1024; // 256 MB
```

- [ ] **Step 4: Run the test**

```bash
npx jest tests/builder.test.ts --testNamePattern="backpressureThreshold" --no-coverage
```

Expected: PASS.

- [ ] **Step 5: Run all tests**

```bash
npx jest --no-coverage
```

Expected: all pass.

- [ ] **Step 6: Commit**

```bash
git add src/index.ts tests/builder.test.ts
git commit -m "fix: lower backpressureThreshold default from 512MB to 256MB"
```

---

## Task 9: Update README.md

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Add ZIP streaming section after the existing multi-file section**

Find the section that documents `.file()` in README.md and add after it:

````markdown
## ZIP Streaming

When `.file()` produces multiple `.xlsx` files, use `.asZip()` to deliver them as a single ZIP archive. This is the only way to stream multi-file output via `.pipe()` or `.toBuffer()`.

### Stream ZIP to HTTP response

```js
// Express / Fastify
res.setHeader('Content-Type', 'application/zip');
res.setHeader('Content-Disposition', 'attachment; filename="export.zip"');

await OracleSqlToExcel()
  .connectionFactory(() => pool.getConnection())
  .file('report', (f) => f
    .maxRowsPerFile(1_000_000)
    .sheet('Data', (s) => s
      .sql('SELECT * FROM BIG_TABLE')
      .columns(COLS)
      .maxRowsPerSheet(500_000)
    )
  )
  .asZip()
  .pipe(res); // streams ZIP directly — no temp file
```

### Write ZIP to disk

```js
const result = await OracleSqlToExcel()
  .connectionFactory(() => pool.getConnection())
  .outputDir('/tmp')
  .filePrefix('export-2026')      // ZIP written as /tmp/export-2026.zip
  .file('report', (f) => f
    .maxRowsPerFile(1_000_000)
    .sheet('Data', (s) => s.sql(SQL).columns(COLS))
  )
  .asZip()
  .run();

console.log(result.file); // /tmp/export-2026.zip
```

### ZIP entry names

Entries inside the ZIP use sequential naming:

```
export-2026.zip
  ├─ report_1.xlsx   ← rows 1–1,000,000
  ├─ report_2.xlsx   ← rows 1,000,001–2,000,000
  └─ report_3.xlsx   ← rows 2,000,001–2,700,000
```

`.asZip()` is only effective when `.file()` is also used. Without `.file()`, it is ignored and output is a plain `.xlsx`.

---

## Memory guide

The library streams rows from Oracle in batches and writes them through ExcelJS to the output. Memory stays bounded by the batch size, not the total row count — **for most use cases**.

### What controls memory

| Setting | Default | Effect |
|---------|---------|--------|
| `.fetchSize(n)` (per sheet) | 50,000 | Rows fetched per Oracle round-trip. Reduce for very wide rows. |
| `.backpressureThreshold(bytes)` | 256 MB | RSS threshold — Oracle fetch pauses when exceeded. |
| `.maxRowsPerFile(n)` (per file) | unlimited | Bounds the size of each XLSX file processed at once. |

### When OOM can still occur

RSS polling is **reactive** — it checks after each batch, not during. If a single batch generates more RSS than available memory, the process will OOM before the pause fires.

**Mitigation for very wide rows or limited server RAM:**

```js
OracleSqlToExcel()
  .backpressureThreshold(128 * 1024 * 1024) // 128 MB — pause earlier
  .file('report', (f) => f
    .maxRowsPerFile(500_000)                  // smaller files = lower peak
    .sheet('Data', (s) => s
      .fetchSize(10_000)                      // smaller batches for wide rows
      .sql(SQL).columns(COLS)
    )
  )
  .asZip()
  .pipe(res);
```

### `.toBuffer()` warning

`.toBuffer()` holds the **entire ZIP in memory** before returning. For large exports this will OOM. Use `.run()` (disk) or `.pipe()` (stream) instead.
````

- [ ] **Step 2: Commit**

```bash
git add README.md
git commit -m "docs: add ZIP streaming section and memory guide to README"
```

---

## Task 10: Update CHANGELOG.md

**Files:**
- Modify: `CHANGELOG.md`

- [ ] **Step 1: Add v2.0.0 entry at the top of CHANGELOG.md**

```markdown
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
```

- [ ] **Step 2: Commit**

```bash
git add CHANGELOG.md
git commit -m "docs: add v2.0.0 CHANGELOG entry"
```

---

## Task 11: Build + full verification

**Files:** none new

- [ ] **Step 1: Run full test suite with coverage**

```bash
npx jest --coverage
```

Expected: all tests pass. Coverage report generated in `coverage/`.

- [ ] **Step 2: Build TypeScript**

```bash
npm run build
```

Expected: no TypeScript errors. `dist/index.js` and `dist/index.d.ts` regenerated.

- [ ] **Step 3: Verify `ZipRunResult` is exported from built types**

```bash
grep "ZipRunResult" dist/index.d.ts
```

Expected: `export interface ZipRunResult extends Result { file: string; }` present.

- [ ] **Step 4: Verify `asZip` method is in built types**

```bash
grep "asZip" dist/index.d.ts
```

Expected: `asZip(value?: boolean): this;` present.

- [ ] **Step 5: Smoke test the built output**

```bash
node -e "
const { OracleSqlToExcel } = require('./dist/index.js');
const b = OracleSqlToExcel().asZip();
console.log('asZip ok:', typeof b.asZip);
console.log('backpressureThreshold default:', b._backpressureThreshold === 256 * 1024 * 1024 ? '256MB OK' : 'WRONG: ' + b._backpressureThreshold);
"
```

Expected:
```
asZip ok: function
backpressureThreshold default: 256MB OK
```

- [ ] **Step 6: Update package.json version to 2.0.0**

In `package.json`, change:
```json
"version": "1.3.0",
```
to:
```json
"version": "2.0.0",
```

- [ ] **Step 7: Final commit**

```bash
git add package.json dist/
git commit -m "release: v2.0.0 — ZIP streaming, showTotalRows fix, 256MB backpressure default"
```

---

## Self-review checklist

- [x] **Spec coverage:** `.asZip()` ✓, `ZipRunResult` ✓, `_executeAsZip` ✓, globalRowOffset fix ✓, `backpressureThreshold` 256MB ✓, `.pipe()` error message ✓, `.run()` ZIP path ✓, `.toBuffer()` ZIP path ✓, README ZIP section ✓, README memory guide ✓, CHANGELOG ✓
- [x] **No placeholders:** all steps have actual code
- [x] **Type consistency:** `ZipRunResult` defined in Task 2, used in Task 6 `.run()` return type. `_executeAsZip(outputStream: Writable, drainFn)` defined in Task 4, called in Tasks 4/5/6. `globalRowOffset: number = 0` parameter added in Task 3, used in all callers.
- [x] **Backward compat tested:** Task 6 includes explicit test that `.run()` without `.asZip()` still returns `MultiRunResult` with `.xlsx` files.
- [x] **drainFn note:** `_executeAsZip` uses RSS-only drainFn for both `.pipe()` and `.run()` paths. Event-driven drain (write() interception) was omitted intentionally — behind nginx/reverse proxy it never fires anyway, and RSS polling handles the same scenario. Can be added in v2.1 if needed.
- [x] **ZIP entry naming:** Always sequential (`{name}_1.xlsx`, `{name}_2.xlsx`). Row-range naming is impossible in streaming mode because the end row is unknown when the ZIP entry is opened. The showTotalRows text inside each XLSX already shows the correct range.
