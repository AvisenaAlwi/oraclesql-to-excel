# oraclesql-to-excel JS

Stream Oracle SQL query results directly into Excel (`.xlsx`) or CSV files with a fluent, chainable API. Built for exports with millions of rows — memory stays bounded regardless of result set size.

- **Streaming** — rows piped from Oracle directly into the output; no full dataset held in memory
- **Excel & CSV** — choose `.xlsx` for rich formatting or `.csv` for maximum throughput and compatibility
- **Multi-sheet** — multiple SQL queries, each on its own sheet, in one workbook
- **Split sheets** — automatically create new sheets when a row limit is reached
- **Multi-file** — split one export across multiple `.xlsx` files via `.file()` + `.maxRowsPerFile()`
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
const { OracleSqlToExcel, OracleSqlToCsv } = require('@avisenaalwi/oraclesql-to-excel');
```

```js
// ESM
import { OracleSqlToExcel, OracleSqlToCsv } from '@avisenaalwi/oraclesql-to-excel';
```

```ts
// TypeScript — full type inference, no extra @types needed
import { OracleSqlToExcel, OracleSqlToCsv } from '@avisenaalwi/oraclesql-to-excel';
import type {
  ColumnDef, HeaderStyle, DocHeaderRow,
  RunResult, MultiRunResult, BufferResult,
  CsvResult, CsvRunResult, CsvBufferResult,
} from '@avisenaalwi/oraclesql-to-excel';
```

---

## Quick Start

### Excel

```js
const { OracleSqlToExcel } = require('@avisenaalwi/oraclesql-to-excel');

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

### CSV

```js
const { OracleSqlToCsv } = require('@avisenaalwi/oraclesql-to-excel');

// Write to file
const { success, file, rowsWritten } = await OracleSqlToCsv()
  .connectionFactory(() => pool.getConnection())
  .sql('SELECT CODE, NAME, BALANCE FROM ACCOUNTS')
  .columns([
    { key: 'CODE',    header: 'Code'    },
    { key: 'NAME',    header: 'Name'    },
    { key: 'BALANCE', header: 'Balance' },
  ])
  .run('/var/reports/accounts.csv');

console.log(success, file, rowsWritten);
```

---

## API Reference

### `OracleSqlToExcel()`

Returns a new `OracleSqlToExcelBuilder`. All methods are chainable.

#### Workbook-level methods

| Method | Default | Description |
|--------|---------|-------------|
| `.connectionFactory(fn)` | — | **Required.** `fn` must return `Promise<Connection>`. Called once per connection needed. |
| `.executeOptions(obj)` | `{}` | Default Oracle execute options for all sheets. |
| `.outputDir(path)` | `process.cwd()` | Directory for `.run()` output. |
| `.filePrefix(name)` | `'export'` | Output filename without extension. Saved as `<name>.xlsx`. |
| `.compress(bool, level?)` | `false` / `1` | Enable XLSX ZIP compression (and outer ZIP level when using `.asZip()`). `level` is zlib `0`–`9`, default `1`. Slower but smaller file. Recommended for `.run()`, not `.pipe()`. |
| `.debug(bool)` | `false` | Verbose logging. Active only when `NODE_ENV` is not `production`. |
| `.onProgress(cb)` | — | Called after each fetch batch. See [Progress Tracking](#progress-tracking-websocket--sse). |
| `.backpressureThreshold(bytes)` | `268435456` (256 MB) | Pause Oracle fetch when process RSS exceeds this value during `.pipe()`. See [Backpressure & Memory](#backpressure--memory). |
| `.sheet(name, fn)` | — | Add a sheet. `name` can be a string or array of strings. |
| `.file(name, fn)` | — | Add a named output file. Only supported with `.run()`. See [Multi-file Export](#multi-file-export). |

#### Terminal methods

| Method | Returns | Description |
|--------|---------|-------------|
| `.run()` | `Promise<RunResult \| MultiRunResult>` | Write workbook to disk. Returns `MultiRunResult` when `.file()` is used. |
| `.pipe(stream)` | `Promise<Result>` | Stream workbook to any `Writable`. |
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

### `FileConfig` — per-file methods

Received as argument `f` inside the `.file(name, fn)` callback. Only applicable with `.run()`.

| Method | Default | Description |
|--------|---------|-------------|
| `.maxRowsPerFile(n)` | `0` (no split) | Split into multiple `.xlsx` files when data rows exceed `n`. |
| `.sheet(name, fn)` | — | Add a sheet to this file. Same API as the builder's `.sheet()`. |

---

### `OracleSqlToCsv()`

Returns a new `OracleSqlToCsvBuilder`. All methods are chainable.

Unlike Excel, CSV writes each row directly to the output stream with no intermediate archiver or ZIP buffer. Memory usage is `O(fetchSize × row_size)` at all times, making it suitable for any dataset size.

#### Methods

| Method | Default | Description |
|--------|---------|-------------|
| `.connectionFactory(fn)` | — | **Required.** `fn` must return `Promise<Connection>`. |
| `.sql(query, params, opts)` | — | **Required.** Oracle SQL with optional bind parameters and execute options. |
| `.columns(defs)` | auto-detect | Array of `{ key, header }`. Omit to auto-detect from Oracle metadata. |
| `.fetchSize(n)` | `50_000` | Rows fetched per Oracle round-trip. |
| `.separator(char)` | `','` | Field separator. Use `';'` for European Excel, `'\t'` for TSV. |
| `.withBom(bool)` | `true` | Prepend UTF-8 BOM so Windows Excel opens the file with correct encoding. |
| `.onProgress(cb)` | — | Called after each fetch batch with `{ rowsWritten }`. |

#### Terminal methods

| Method | Returns | Description |
|--------|---------|-------------|
| `.run(filepath)` | `Promise<CsvRunResult>` | Write CSV to file at `filepath` (absolute or relative, including extension). |
| `.pipe(stream)` | `Promise<CsvResult>` | Stream CSV to any `Writable` (e.g. Express `res`). |
| `.toBuffer()` | `Promise<CsvBufferResult>` | Return entire CSV as `Buffer`. |

---

### `ColumnDef`

```js
{
  key       : 'COLUMN_NAME',          // Oracle column name — case-sensitive
  header    : 'Display Label',        // Header text. Defaults to key.
  type      : 'text',                 // 'text' | 'number' | 'date' | 'datetime'  (Excel only)
  width     : 18,                     // Column width in Excel character units     (Excel only)
  numFmt    : '#,##0.00',             // Excel number format string                (Excel only)
  align     : 'right',                // 'left' | 'center' | 'right'              (Excel only)
  wrapText  : false,                  // Enable text wrap                          (Excel only)
  bgColor   : 'FFFF00',               // Cell background color (hex)               (Excel only)
  fontColor : 'FF0000',               // Cell font color (hex)                     (Excel only)
}
```

For CSV, only `key` and `header` are used.

**Type notes (Excel):**
- `number` — values with > 15 significant digits are cast to string to prevent Excel precision loss.
- `date` — default format `dd/mm/yyyy`.
- `datetime` — default format `dd/mm/yyyy hh:mm:ss`.
- `text` — all values cast via `String()`.

---

### Return Values

#### `OracleSqlToExcel().run()` → `RunResult`

```js
{
  success     : true,
  file        : '/output/report.xlsx',
  sheets      : ['Sheet1', 'Sheet1 2', 'Sheet2'],
  skippedRows : 0,
  error       : undefined,   // present only on failure
}
```

#### `OracleSqlToExcel().run()` with `.file()` → `MultiRunResult`

```js
{
  success     : true,
  files       : [
    {
      file  : '/output/data_1.xlsx',
      sheets: [{ name: 'Data', startRow: 1, endRow: 1000000 }],
    },
    {
      file  : '/output/data_2.xlsx',
      sheets: [{ name: 'Data', startRow: 1000001, endRow: 1823400 }],
    },
  ],
  skippedRows : 0,
  error       : undefined,
}
```

#### `.pipe()` → `Result`

```js
{ success, sheets, skippedRows, error? }
```

#### `.toBuffer()` → `BufferResult`

```js
{ success, buffer, sheets, skippedRows, error? }
```

#### `OracleSqlToCsv().run()` → `CsvRunResult`

```js
{ success, file, rowsWritten, error? }
```

#### `OracleSqlToCsv().pipe()` → `CsvResult`

```js
{ success, rowsWritten, error? }
```

#### `OracleSqlToCsv().toBuffer()` → `CsvBufferResult`

```js
{ success, buffer, rowsWritten, error? }
```

---

## Examples

### Excel — Write to File

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

### Excel — Stream to HTTP Response

```js
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

### Excel — Return as Buffer (S3 / Email)

```js
const { success, buffer, error } = await OracleSqlToExcel()
  .connectionFactory(() => pool.getConnection())
  .sheet('Report', s => s.sql('SELECT * FROM MONTHLY_REPORT').columns(COLS))
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

### Excel — Multi-Sheet Workbook

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
  .run();

console.log('Sheets written:', result.sheets);
// → ['Summary', 'Detail', 'Detail 2', ...]
```

---

### Multi-file Export

Split one large result across multiple `.xlsx` files. Oracle's `ResultSet` stays open across files — only one query per sheet regardless of how many files are produced.

```js
const { files } = await OracleSqlToExcel()
  .connectionFactory(() => pool.getConnection())
  .outputDir('/var/reports')
  .file('transactions', f => f
    .maxRowsPerFile(1_000_000)   // one .xlsx per million rows
    .sheet('Data', s => s
      .sql('SELECT * FROM TRANSACTIONS WHERE YEAR = :y', { y: 2026 })
      .columns(COLS)
      .maxRowsPerSheet(900_000)  // split sheet within each file too
    )
  )
  .run();

// files → [
//   { file: '/var/reports/transactions_1.xlsx' },
//   { file: '/var/reports/transactions_2.xlsx' },
// ]
console.log(`${files.length} file(s) written`);
```

> **Cross-file navigation notes** — when a query spans multiple files, the library automatically inserts informational rows at file boundaries so the reader knows where data continues:
> - **Start of file 2+** (before "Showing rows…" and column headers): `"Previous data on file: transactions_1.xlsx"` — italic, gray
> - **End of files 1 to N-1** (after last data row): `"Next data available on file: transactions_2.xlsx"` — italic, gray
>
> These use the same visual style as the within-file sheet continuation notes (`"Continued on sheet: …"`).

> **`maxRowsPerSheet` and `maxRowsPerFile` interaction** — if `maxRowsPerSheet` is set larger than `maxRowsPerFile`, it is automatically capped at the file limit. Each file segment will never split into more sheets than needed.

Multiple `.file()` calls produce independent output files:

```js
await OracleSqlToExcel()
  .connectionFactory(() => pool.getConnection())
  .outputDir('/var/reports')
  .file('summary', f => f.sheet('Summary', s => s.sql(SQL1).columns(COLS1)))
  .file('detail',  f => f.sheet('Detail',  s => s.sql(SQL2).columns(COLS2)))
  .run();
// → /var/reports/summary.xlsx, /var/reports/detail.xlsx
```

> `.file()` is only supported with `.run()`. Using it with `.pipe()` or `.toBuffer()` throws an error.

---

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

---

### CSV — Write to File

```js
const { OracleSqlToCsv } = require('@avisenaalwi/oraclesql-to-excel');

const { success, file, rowsWritten } = await OracleSqlToCsv()
  .connectionFactory(() => pool.getConnection())
  .sql('SELECT CODE, NAME, AMOUNT FROM BIG_TABLE')
  .columns([
    { key: 'CODE',   header: 'Code'   },
    { key: 'NAME',   header: 'Name'   },
    { key: 'AMOUNT', header: 'Amount' },
  ])
  .run('/var/reports/export.csv');

console.log(`${rowsWritten} rows → ${file}`);
```

---

### CSV — Stream to HTTP Response

Memory stays `O(fetchSize × row_size)` regardless of total row count — no backpressure issues even for 10M+ rows.

```js
app.get('/download/csv', async (req, res) => {
  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', 'attachment; filename="report.csv"');

  const result = await OracleSqlToCsv()
    .connectionFactory(() => pool.getConnection())
    .sql('SELECT CODE, NAME, AMOUNT FROM BIG_TABLE WHERE PERIOD = :p', { p: req.query.period })
    .columns([
      { key: 'CODE',   header: 'Code'   },
      { key: 'NAME',   header: 'Name'   },
      { key: 'AMOUNT', header: 'Amount' },
    ])
    .pipe(res);

  if (!result.success) {
    console.error('CSV export failed mid-stream:', result.error);
  }
});
```

---

### CSV — Custom Separator

```js
// Semicolon — European Excel (avoids conflict with decimal comma)
OracleSqlToCsv().separator(';')

// Tab-separated values (TSV)
OracleSqlToCsv().separator('\t')

// No BOM — for non-Windows consumers
OracleSqlToCsv().withBom(false)
```

---

### Bind Parameters

```js
// Named binds (recommended for Oracle)
s.sql(
  'SELECT * FROM TRANSACTIONS WHERE BRANCH_CODE = :branch AND PERIOD = :period',
  { branch: '019', period: '202601' }
)

// Multiple conditions with date binds
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

**Excel:**

```js
await OracleSqlToExcel()
  .connectionFactory(() => pool.getConnection())
  .onProgress(({ sheet, rowsWritten, skippedRows, totalRowsWritten }) => {
    io.emit('export-progress', { sheet, rowsWritten, skippedRows, totalRowsWritten });
  })
  .sheet('Data', s => s.sql(SQL).columns(COLS))
  .run();
```

**CSV:**

```js
await OracleSqlToCsv()
  .connectionFactory(() => pool.getConnection())
  .sql(SQL).columns(COLS)
  .onProgress(({ rowsWritten }) => {
    io.emit('export-progress', { rowsWritten });
  })
  .run('/tmp/export.csv');
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

### Auto-detect Columns

Omit `.columns()` to auto-detect from Oracle metadata. Column names and types are inferred automatically.

```js
// Excel
s.sql('SELECT CODE, NAME, BALANCE FROM ACCOUNTS')
// → all 3 columns, Oracle names as headers, types inferred from DB metadata

// CSV
OracleSqlToCsv().sql('SELECT CODE, NAME, BALANCE FROM ACCOUNTS')
// → all 3 columns, Oracle names as headers
```

> A warning is emitted in non-production environments. Column order depends on the SELECT clause. Use `.columns()` for stable, long-term output.

---

### Backpressure & Memory

**Excel (`.pipe()`):** After each Oracle fetch batch, the library checks process RSS. If it exceeds `.backpressureThreshold()` (default: 256 MB), the Oracle fetch pauses and polls every 200 ms until RSS drops. This handles the common case where Node.js sits behind a reverse proxy (nginx, etc.) that accepts data instantly — `stream.write()` always returns `true`, yet data accumulates in Node.js output buffers.

> `.backpressureThreshold()` has no effect on `.run()` (file writes drain naturally) or `.toBuffer()` (data collected in memory by design).

#### Why the stream appears to freeze during a pause

The pause happens **between** Oracle batches — after ExcelJS finishes writing a batch but before the next `getRows()` call. During the pause:

- No new rows → ExcelJS produces no new XML → no new bytes flow to the browser
- Data already flushed **before** the pause was already received by the browser — only the production of new data stops

In a browser's Network DevTools you will see the download size counter stall for a few seconds, then resume. This is expected and correct behavior.

#### RSS stays near the threshold — why?

Each batch follows this cycle:

```
getRows(fetchSize) → ExcelJS writes XML → RSS rises
                                         ↓
                              RSS > threshold → pause
                                         ↓
                              GC recovers some RAM
                              (active ZIP entry stays open → not all RAM freed)
                                         ↓
                              RSS < threshold → resume → next batch
```

RSS does not drop all the way back to baseline because the active ExcelJS workbook and open archiver ZIP entry hold live buffers. GC only reclaims the overhead from the previous batch. The result is that RSS oscillates just below and above the threshold for the entire export — this is normal.

#### Estimating RSS per batch

Each batch costs roughly:

```
RSS_per_batch ≈ fetchSize × avg_bytes_per_row × 3
```

The `× 3` factor accounts for ExcelJS XML expansion (raw data → XML tags) plus archiver ZIP buffering. This is a rough estimate — actual values vary by column count, data types, and string length.

| Row width | `fetchSize 50 000` | `fetchSize 10 000` |
|-----------|-------------------|-------------------|
| Narrow — ~100 B/row (few short columns) | ~15 MB | ~3 MB |
| Average — ~300 B/row (10–15 mixed columns) | ~45 MB | ~9 MB |
| Wide — ~1 KB/row (many columns or long strings) | ~150 MB | ~30 MB |

**Rule of thumb for setting `backpressureThreshold`:**

```
threshold = baseline_RSS + (2 × RSS_per_batch)
```

- `baseline_RSS` — process RSS before any export starts (check with `process.memoryUsage().rss`)
- Leaving headroom of `2 × RSS_per_batch` ensures one full batch can be processed before the pause triggers

#### Tuning example — 2.7 M rows, average row width

```js
// Baseline RSS ~80 MB, average row ~300 B, fetchSize 10 000 → ~9 MB/batch
// threshold = 80 + (2 × 9) ≈ 100 MB
OracleSqlToExcel()
  .fetchSize(10_000)                          // smaller batches → smaller RSS spike
  .backpressureThreshold(100 * 1024 * 1024)  // pause earlier → GC more effective
  .file('report', f => f
    .sheet('Data', s => s.sql(SQL).columns(COLS))
  )
  .asZip()
  .pipe(res);
```

Smaller `fetchSize` is the most effective lever: it directly reduces RSS spike per batch and makes pauses shorter and less visible to the end user.

**CSV (`.pipe()`):** CSV rows are plain text — far smaller than Excel's XML+ZIP output. At 50,000 rows/batch, a typical batch is a few MB of text and flushes quickly through the TCP stack. Backpressure is generally not a concern for CSV streams.

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

## Excel vs CSV — When to Use Which

| | Excel (`.xlsx`) | CSV |
|---|---|---|
| **Formatting** | Number formats, colors, fonts, freeze header, auto-filter | None |
| **Multiple sheets** | Yes | No |
| **Max rows per sheet** | 1,048,576 (auto-split) | Unlimited |
| **File size** | Larger (XML inside ZIP) | Smaller (~3–10× smaller) |
| **Memory (stream)** | Bounded via RSS polling | `O(fetchSize × row_size)` — minimal |
| **Speed** | Slower (XML generation + ZIP) | Faster |
| **Best for** | Formatted reports, pivot tables | Raw data exports, BI tools, large datasets |

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
