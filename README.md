# @avisenaalwi/oraclesql-to-excel

Stream Oracle SQL query results directly into Excel (`.xlsx`) files with a fluent, chainable API.

- **Streaming** — rows piped from Oracle directly into ExcelJS; no full dataset held in memory
- **Multi-sheet** — multiple SQL queries, each on its own sheet, in one workbook
- **Split sheets** — automatically create new sheets when a row limit is reached
- **HTTP streaming** — pipe directly to an Express/Fastify response, no temp file
- **Buffer output** — return as `Buffer` for S3 uploads, email attachments, or DB BLOBs
- **Document headers** — custom rows above the table (title, period, logo placeholders) with merge and style support
- **Row range summary** — "Showing rows X – Y of Z total" auto-prepended per sheet via a parallel COUNT query
- **Auto-filter & freeze header** — one method call each
- **Per-column formatting** — number formats, alignment, wrap text, font/bg colors

---

## Installation

```bash
npm install @avisenaalwi/oraclesql-to-excel
```

`oracledb` is a peer dependency — install it separately if not already present:

```bash
npm install oracledb
```

---

## Import Styles

```js
// CommonJS
const { OracleSqlToExcel } = require('@avisenaalwi/oraclesql-to-excel');
```

```js
// ESM
import { OracleSqlToExcel } from '@avisenaalwi/oraclesql-to-excel';
```

```ts
// TypeScript — full type inference, no extra @types needed
import { OracleSqlToExcel } from '@avisenaalwi/oraclesql-to-excel';
import type { ColumnDef, HeaderStyle, DocHeaderRow, RunResult, BufferResult } from '@avisenaalwi/oraclesql-to-excel';
```

---

## Quick Start

**JavaScript (CommonJS)**

```js
const oracledb       = require('oracledb');
const { OracleSqlToExcel } = require('@avisenaalwi/oraclesql-to-excel');

const pool = await oracledb.createPool({ /* your pool config */ });

const { success, file } = await OracleSqlToExcel()
  .connectionFactory(() => pool.getConnection())
  .sheet('Report', s => s
    .sql('SELECT CODE, NAME, BALANCE FROM ACCOUNTS')
    .columns([
      { key: 'CODE',    header: 'Code',    type: 'text',   width: 12 },
      { key: 'NAME',    header: 'Name',    type: 'text',   width: 30 },
      { key: 'BALANCE', header: 'Balance', type: 'number', width: 18, numFmt: '#,##0.00' },
    ])
    .freezeHeader()
    .autoFilter()
    .headerStyle({ bgColor: '4472C4', fontColor: 'FFFFFF' })
  )
  .filePrefix('account-report')
  .run();

console.log(success, file); // true, '/path/to/account-report.xlsx'
```

**TypeScript**

```ts
import oracledb                              from 'oracledb';
import { OracleSqlToExcel }                        from '@avisenaalwi/oraclesql-to-excel';
import type { ColumnDef, RunResult }         from '@avisenaalwi/oraclesql-to-excel';

const pool = await oracledb.createPool({ /* your pool config */ });

const COLS: ColumnDef[] = [
  { key: 'CODE',    header: 'Code',    type: 'text',   width: 12 },
  { key: 'NAME',    header: 'Name',    type: 'text',   width: 30 },
  { key: 'BALANCE', header: 'Balance', type: 'number', width: 18, numFmt: '#,##0.00' },
];

const result: RunResult = await OracleSqlToExcel()
  .connectionFactory(() => pool.getConnection())
  .sheet('Report', s => s
    .sql('SELECT CODE, NAME, BALANCE FROM ACCOUNTS')
    .columns(COLS)
    .freezeHeader()
    .autoFilter()
    .headerStyle({ bgColor: '4472C4', fontColor: 'FFFFFF' })
  )
  .filePrefix('account-report')
  .run();

console.log(result.success, result.file);
```

---

## API Reference

### `OracleSqlToExcel()`

Returns a new `OracleSqlToExcelBuilder` instance. All methods are chainable.

#### Workbook-level methods

| Method | Default | Description |
|--------|---------|-------------|
| `.connectionFactory(fn)` | — | **Required.** `fn` must return `Promise<Connection>`. Called once per connection needed. |
| `.executeOptions(obj)` | `{}` | Default Oracle execute options for all sheets. |
| `.outputDir(path)` | `process.cwd()` | Directory for `.run()` output. |
| `.filePrefix(name)` | `'export'` | Output filename without extension. Saved as `<name>.xlsx`. |
| `.compress(bool)` | `false` | Enable XLSX ZIP compression. Slower but smaller file. Recommended for `.run()`, not `.pipe()`. |
| `.debug(bool)` | `false` | Verbose logging. Active only when `NODE_ENV` is not `production`. |
| `.onProgress(cb)` | — | Called after each fetch batch. See [Progress Tracking](#progress-tracking-websocket--sse). |
| `.backpressureThreshold(bytes)` | `16777216` (16 MB) | Pause Oracle fetch when the output stream buffer exceeds this size. Only applies to `.pipe()`. See [Backpressure & Memory](#backpressure--memory-pipe-only). |
| `.sheet(name, fn)` | — | Add a sheet. `name` can be a string or array of strings. |

#### Terminal methods

| Method | Returns | Description |
|--------|---------|-------------|
| `.run()` | `Promise<RunResult>` | Write workbook to disk. Deletes partial file on error. |
| `.pipe(stream)` | `Promise<Result>` | Stream workbook to any `WritableStream`. |
| `.toBuffer()` | `Promise<BufferResult>` | Return workbook as `Buffer`. |

---

### `SheetConfig` — per-sheet methods

Received as argument `s` inside the `.sheet(name, fn)` callback.

| Method | Default | Description |
|--------|---------|-------------|
| `.sql(query, params, opts)` | — | **Required.** Oracle SQL, optional bind params and execute option overrides. |
| `.columns(defs)` | auto-detect | Array of `ColumnDef`. Omit to auto-detect from Oracle metadata. |
| `.maxRowsPerSheet(n)` | `1_000_000` | Max data rows before a new physical sheet is created. |
| `.fetchSize(n)` | `50_000` | Rows fetched per Oracle round-trip. |
| `.freezeHeader()` | off | Freeze the header row. |
| `.autoFilter()` | off | Add dropdown filter to every header column. |
| `.headerStyle(obj)` | bold | Override column header row style. See `HeaderStyle`. |
| `.docHeader(rows)` | none | Custom rows above the table header. First sheet only. See [Document Header](#document-header). |
| `.showTotalRows()` | off | Prepend "Showing rows X – Y of Z total". Runs COUNT in parallel. |
| `.onRowError(mode)` | `'throw'` | `'throw'` aborts on first bad row. `'skip'` drops it and continues. |

---

### `ColumnDef`

```js
{
  key       : 'COLUMN_NAME',          // Oracle column name — case-sensitive
  header    : 'Display Label',        // Header text. Defaults to key.
  type      : 'text',                 // 'text' | 'number' | 'date' | 'datetime'
  width     : 18,                     // Column width in Excel character units
  numFmt    : '#,##0.00',             // Excel number format string
  align     : 'right',                // 'left' | 'center' | 'right'
  wrapText  : false,                  // Enable text wrap
  bgColor   : 'FFFF00',               // Cell background color (hex, with or without #)
  fontColor : 'FF0000',               // Cell font color (hex)
}
```

**Type notes:**
- `number` — values with > 15 significant digits are cast to string to prevent Excel precision loss.
- `date` — default format `dd/mm/yyyy`.
- `datetime` — default format `dd/mm/yyyy hh:mm:ss`.
- `text` — all values cast via `String()`.

---

### `HeaderStyle`

```js
{
  bold      : true,     // Default: true
  bgColor   : '4472C4', // Background color hex
  fontColor : 'FFFFFF', // Font color hex
}
```

---

### Return Values

#### `.run()` → `RunResult`

```js
{
  success     : true,
  file        : '/output/report.xlsx',
  sheets      : ['Sheet1', 'Sheet1 2', 'Sheet2'],
  skippedRows : 0,
  error       : undefined,   // present only on failure
}
```

#### `.pipe()` → `Result`

```js
{ success, sheets, skippedRows, error? }
```

#### `.toBuffer()` → `BufferResult`

```js
{
  success     : true,
  buffer      : Buffer,   // empty Buffer when success=false
  sheets      : [...],
  skippedRows : 0,
  error       : undefined,
}
```

---

## Examples

### Write to File

```js
const { OracleSqlToExcel } = require('@avisenaalwi/oraclesql-to-excel');

const result = await OracleSqlToExcel()
  .connectionFactory(() => pool.getConnection())
  .outputDir('/var/reports')
  .filePrefix('monthly-report')
  .sheet('Summary', s => s
    .sql('SELECT * FROM SUMMARY_VIEW')
    .columns([
      { key: 'PERIOD', header: 'Period', type: 'text',   width: 15 },
      { key: 'TOTAL',  header: 'Total',  type: 'number', width: 20, numFmt: '#,##0' },
    ])
    .freezeHeader()
    .autoFilter()
  )
  .compress(true)   // smaller file for disk storage
  .run();

if (!result.success) {
  console.error('Export failed:', result.error);
} else {
  console.log('Saved to:', result.file);
}
```

---

### Stream to Express HTTP Response

```js
const express        = require('express');
const { OracleSqlToExcel } = require('@avisenaalwi/oraclesql-to-excel');

const app = express();

app.get('/download/report', async (req, res) => {
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="report.xlsx"');

  const result = await OracleSqlToExcel()
    .connectionFactory(() => pool.getConnection())
    .sheet('Data', s => s
      .sql(
        'SELECT CODE, NAME, BALANCE FROM ACCOUNTS WHERE PERIOD = :period',
        { period: req.query.period }
      )
      .columns([
        { key: 'CODE',    header: 'Code',    type: 'text',   width: 12 },
        { key: 'NAME',    header: 'Name',    type: 'text',   width: 30 },
        { key: 'BALANCE', header: 'Balance', type: 'number', width: 18, numFmt: '#,##0.00' },
      ])
      .freezeHeader()
    )
    // compress(false) recommended for HTTP — web server handles gzip separately
    .pipe(res);

  if (!result.success) {
    // Headers already sent — cannot send error response, log instead
    console.error('Export failed mid-stream:', result.error);
  }
});
```

---

### Return as Buffer (S3 / Email)

```js
const AWS            = require('aws-sdk');
const { OracleSqlToExcel } = require('@avisenaalwi/oraclesql-to-excel');

const s3 = new AWS.S3();

const { success, buffer, error } = await OracleSqlToExcel()
  .connectionFactory(() => pool.getConnection())
  .sheet('Report', s => s
    .sql('SELECT * FROM MONTHLY_REPORT')
    .columns(COLS)
  )
  .toBuffer();

if (!success) throw new Error(error);

await s3.putObject({
  Bucket     : 'my-bucket',
  Key        : 'reports/monthly.xlsx',
  Body       : buffer,
  ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
}).promise();
```

---

### Multi-Sheet Workbook

```js
const result = await OracleSqlToExcel()
  .connectionFactory(() => pool.getConnection())
  .filePrefix('full-report')
  .sheet('Summary', s => s
    .sql('SELECT * FROM V_MONTHLY_SUMMARY WHERE YEAR = :year', { year: 2026 })
    .columns(COLS_SUMMARY)
    .freezeHeader()
    .autoFilter()
    .headerStyle({ bgColor: '4472C4', fontColor: 'FFFFFF' })
  )
  .sheet('Detail', s => s
    .sql('SELECT * FROM V_TRANSACTION_DETAIL WHERE YEAR = :year', { year: 2026 })
    .columns(COLS_DETAIL)
    .freezeHeader()
    .onRowError('skip')   // skip bad rows instead of aborting
  )
  .sheet('Account Ref', s => s
    .sql('SELECT CODE, NAME, TYPE FROM MASTER_ACCOUNT ORDER BY CODE')
    .columns(COLS_ACCOUNT)
  )
  .run();

console.log('Sheets written:', result.sheets);
// → ['Summary', 'Detail', 'Detail 2', 'Account Ref']
```

---

### Bind Parameters

```js
// Named binds (recommended for Oracle)
s.sql(
  'SELECT * FROM TRANSACTIONS WHERE BRANCH_CODE = :branch AND PERIOD = :period',
  { branch: '019', period: '202601' }
)

// Multiple conditions
s.sql(
  'SELECT * FROM TRANSACTIONS WHERE STATUS = :status AND TXN_DATE >= :from AND TXN_DATE <= :to',
  { status: 'A', from: new Date('2026-01-01'), to: new Date('2026-01-31') }
)
```

---

### Document Header

Add custom rows above the table header — company name, report title, period, etc.

```js
s.docHeader([
  {
    text  : 'ACME CORPORATION',
    style : { bold: true, fontSize: 14, align: 'center' },
    height: 24,
  },
  {
    text  : 'BANK STATEMENT REPORT',
    style : { bold: true, fontSize: 12, align: 'center' },
    height: 20,
  },
  {
    text  : 'Period: January 2026',
    style : { italic: true, align: 'center' },
  },
  { text: '' },  // spacer row
])
```

**Multi-column doc header** — for logo + title side-by-side:

```js
s.docHeader([
  {
    columns: [
      // Columns 1-2, spans 2 rows downward
      {
        text       : '[LOGO]',
        mergeAcross: 1,
        mergeDown  : 1,
        style      : { align: 'center', bold: true, bgColor: 'E9EFF7' },
      },
      // Columns 3-10, row 1
      {
        text       : 'BANK STATEMENT REPORT',
        mergeAcross: 7,
        style      : { bold: true, fontSize: 14, align: 'center' },
      },
    ],
    height: 30,
  },
  {
    columns: [
      // Columns 1-2 blocked by mergeDown above — skipped automatically
      // Columns 3-10, row 2
      {
        text       : 'Period: January 2026',
        mergeAcross: 7,
        style      : { italic: true, align: 'center' },
      },
    ],
    height: 20,
  },
  { text: '' },  // spacer
])
```

---

### Row Range Summary (`showTotalRows`)

Prepends `"Showing rows 1 – 50,000 of 182,400 total"` on each sheet.
Runs `SELECT COUNT(*)` in parallel — no overhead during the stream itself.

```js
s.sql('SELECT * FROM LARGE_TABLE')
 .showTotalRows()
```

Output example across split sheets:
```
Sheet 1 → "Showing rows 1 – 1,000,000 of 1,800,000 total"
Sheet 2 → "Showing rows 1,000,001 – 1,800,000 of 1,800,000 total"
```

> **Note:** Silently skipped when the SQL uses a CTE (`WITH ... AS`) or the COUNT query fails.

---

### Split Sheets (`maxRowsPerSheet`)

When a result set exceeds the row limit, the library automatically creates new sheets.

```js
s.sql('SELECT * FROM HUGE_TABLE')
 .columns(COLS)
 .maxRowsPerSheet(500_000)   // split every 500k rows
```

Sheet names follow the pattern: `"Sheet"`, `"Sheet 2"`, `"Sheet 3"`, …

Control split names with an array:

```js
.sheet(['Part A', 'Part B'], s => s
  .sql(SQL)
  .maxRowsPerSheet(200_000)
)
// → 'Part A', 'Part B', 'Part B 2', 'Part B 3', ...
```

Each sheet gets a `"Continued from sheet: <prev>"` notice at the top and a `"Continued on sheet: <next>"` footer at the bottom.

---

### Progress Tracking (WebSocket / SSE)

```js
const io = require('socket.io')(server);

await OracleSqlToExcel()
  .connectionFactory(() => pool.getConnection())
  .onProgress(({ sheet, rowsWritten, skippedRows, totalRowsWritten }) => {
    io.emit('export-progress', { sheet, rowsWritten, skippedRows, totalRowsWritten });
  })
  .sheet('Data', s => s.sql(SQL).columns(COLS))
  .run();
```

Callback fires once per Oracle fetch batch (default 50,000 rows).

---

### Error Handling

```js
// 'throw' mode (default) — abort on first bad row
s.onRowError('throw')

// 'skip' mode — drop bad rows, continue export
s.onRowError('skip')
```

Check skipped row count in the result:

```js
const { success, skippedRows } = await OracleSqlToExcel()
  .connectionFactory(() => pool.getConnection())
  .sheet('Data', s => s.sql(SQL).columns(COLS).onRowError('skip'))
  .run();

if (skippedRows > 0) {
  console.warn(`Export complete — ${skippedRows} rows skipped.`);
}
```

---

### Auto-detect Columns (No `.columns()`)

Omit `.columns()` to auto-detect from Oracle metadata. Column names and types are inferred automatically.

```js
s.sql('SELECT CODE, NAME, BALANCE FROM ACCOUNTS')
// No .columns() — all 3 columns written with Oracle names as headers
```

> A warning is emitted in non-production environments. Column order depends on the SELECT clause and may change if the query or schema changes. Use `.columns()` for stable, long-term output.

---

### Custom Execute Options

```js
// Workbook-level default (applies to all sheets)
OracleSqlToExcel()
  .executeOptions({ autoCommit: false })

// Sheet-level override (merged over workbook-level)
s.sql('SELECT * FROM T', {}, { autoCommit: false })
```

> `outFormat`, `resultSet`, and `fetchArraySize` are always overridden internally and cannot be changed.

---

### Backpressure & Memory (`.pipe()` only)

When streaming to an HTTP response with `.pipe(res)`, the Oracle fetch loop runs much faster than a client can download. Without backpressure control, row XML accumulates in Node.js stream buffers — for exports with millions of rows this can easily exhaust process memory.

The library handles this automatically: after each Oracle fetch batch, it checks the output stream's `writableLength`. If it exceeds the threshold, it pauses and waits for the stream to drain before fetching the next batch.

**Default threshold is 16 MB.** Adjust with `.backpressureThreshold()`:

```js
// Lower threshold — pause sooner, less memory pressure, slightly slower throughput
OracleSqlToExcel()
  .connectionFactory(() => pool.getConnection())
  .backpressureThreshold(8 * 1024 * 1024)   // pause at 8 MB
  .sheet('Data', s => s.sql(SQL).columns(COLS))
  .pipe(res);

// Higher threshold — pause less often, higher throughput, more peak memory allowed
OracleSqlToExcel()
  .backpressureThreshold(32 * 1024 * 1024)  // pause at 32 MB
  .sheet('Data', s => s.sql(SQL).columns(COLS))
  .pipe(res);
```

> **Note:** `.backpressureThreshold()` has no effect on `.run()` (file writes drain naturally) or `.toBuffer()` (all data is intentionally collected in memory).

---

## Requirements

| Dependency | Version |
|------------|---------|
| Node.js    | >= 14.0.0 |
| exceljs    | >= 4.x |
| oracledb   | >= 5.x (peer) |

---

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for release history.

---

## License

[MIT](LICENSE) © Avisena Alwi
