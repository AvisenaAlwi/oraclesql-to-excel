import fs from 'fs';
import ExcelJS from 'exceljs';
import path from 'path';
import { PassThrough, Writable } from 'stream';
import archiver from 'archiver';

// oracledb.OUT_FORMAT_OBJECT (numeric value 4002, stable across v5+)
const OUT_FORMAT_OBJECT    = 4002;
const EXCEL_MAX_ROWS       = 1_000_000;
const EXCEL_MAX_SIG_DIGITS = 15;

// ── Public types ──────────────────────────────────────────────────────────────

/**
 * Style for a single doc-header cell. Used by {@link DocHeaderCell} and simple-mode {@link DocHeaderRow}.
 */
export interface DocHeaderStyle {
  bold?      : boolean;
  italic?    : boolean;
  fontSize?  : number;
  /** Font color hex, e.g. `'FF0000'` or `'#FF0000'`. */
  fontColor? : string;
  /** Background fill color hex, e.g. `'4472C4'` or `'#4472C4'`. */
  bgColor?   : string;
  align?     : 'left' | 'center' | 'right';
}

/**
 * A single cell in a multi-column doc-header row.
 *
 * @example
 * { text: 'LOGO', mergeAcross: 1, mergeDown: 1, style: { align: 'center', bold: true } }
 */
export interface DocHeaderCell {
  text?        : string;
  /** Number of **additional** columns to merge rightward. `mergeAcross: 2` → spans 3 columns. */
  mergeAcross? : number;
  /** Number of **additional** rows to merge downward. `mergeDown: 1` → spans 2 rows. */
  mergeDown?   : number;
  style?       : DocHeaderStyle;
}

/**
 * A single row in the document header block.
 * - **Simple mode**: set `text` (+ optional `merge`, `style`).
 * - **Column mode**: set `columns` — array of {@link DocHeaderCell}.
 */
export interface DocHeaderRow {
  /** Column mode: array of cells in this row. When set, `text` and `merge` are ignored. */
  columns? : DocHeaderCell[];
  /** Simple mode: single text value. Empty string = spacer row. */
  text?    : string;
  /** Simple mode: merge across full column width. Default: `true`. */
  merge?   : boolean;
  /** Row height in points. Applies in both modes. */
  height?  : number;
  /** Simple mode: text style. */
  style?   : DocHeaderStyle;
}

/**
 * Column definition for Excel output.
 */
export interface ColumnDef {
  /** Oracle column name — case-sensitive. */
  key       : string;
  /** Header label in Excel. Defaults to `key`. */
  header?   : string;
  /** Cell type. Numbers exceeding 15 significant digits are coerced to string. Default: `'text'`. */
  type?     : 'text' | 'number' | 'date' | 'datetime';
  /** Column width in Excel character units. Default: `18`. */
  width?    : number;
  /** Custom Excel number format string, e.g. `'#,##0.00'`. Overrides the default for `date`/`datetime`. */
  numFmt?   : string;
  align?    : 'left' | 'center' | 'right';
  wrapText? : boolean;
  /** Background fill color hex, e.g. `'FFFF00'` or `'#FFFF00'`. */
  bgColor?  : string;
  /** Font color hex, e.g. `'FF0000'` or `'#FF0000'`. */
  fontColor?: string;
}

/**
 * Header row style overrides.
 */
export interface HeaderStyle {
  /** Default: `true`. */
  bold?      : boolean;
  /** Background fill color hex, e.g. `'4472C4'` for blue. */
  bgColor?   : string;
  /** Font color hex, e.g. `'FFFFFF'` for white. */
  fontColor? : string;
}

/**
 * Payload delivered to the {@link OracleSqlToExcelBuilder#onProgress} callback after each fetch batch.
 */
export interface ProgressInfo {
  /** Name of the current logical sheet being written. */
  sheet            : string;
  /** Rows successfully written to this sheet so far. */
  rowsWritten      : number;
  /** Rows dropped so far on this sheet (requires `.onRowError('skip')`). */
  skippedRows      : number;
  /** Cumulative rows written across all sheets. */
  totalRowsWritten : number;
}

export interface Result {
  success     : boolean;
  sheets      : string[];
  skippedRows : number;
  error?      : string;
}

export interface FileSegment {
  /** Absolute path to this file. */
  file    : string;
  /** First data row number in this file (1-based, across all files in the split). */
  startRow: number;
  /** Last data row number in this file (inclusive). */
  endRow  : number;
}

export interface RunResult extends Result {
  /** Absolute path to the written `.xlsx` file. Single-file mode only. */
  file: string;
}

export interface MultiRunResult extends Result {
  /** One entry per generated file, in order. */
  files: FileSegment[];
}

export interface BufferResult extends Result {
  /** In-memory workbook buffer. Empty `Buffer` when `success` is `false`. */
  buffer: Buffer;
}

export interface ZipRunResult extends Result {
  /** Absolute path to the written .zip file. */
  file: string;
}

// ── Internal types ────────────────────────────────────────────────────────────

type ColumnType = NonNullable<ColumnDef['type']>;

interface SheetSegmentResult {
  sheetNames  : string[];
  skippedRows : number;
  rowsWritten : number;
  overflowRows: Record<string, unknown>[];
  openRS      : OracleResultSet | null;
}

interface OracleMetaData {
  name        : string;
  dbTypeName? : string;
}

interface OracleResultSet {
  getRows(count: number): Promise<Record<string, unknown>[]>;
  close(): Promise<void>;
}

interface OracleExecuteResult {
  metaData?  : OracleMetaData[];
  resultSet? : OracleResultSet;
  rows?      : Record<string, unknown>[];
}

interface OracleConnection {
  execute(
    sql      : string,
    params?  : Record<string, unknown>,
    options? : Record<string, unknown>
  ): Promise<OracleExecuteResult>;
  close(): Promise<void>;
}

interface WorkbookTarget {
  filename? : string;
  stream?   : Writable;
}

interface ProgressCtx {
  totalRowsWritten: number;
}

type StreamWorksheet = ExcelJS.Worksheet;
type StreamWorkbook  = ExcelJS.stream.xlsx.WorkbookWriter;

// ── Helpers ───────────────────────────────────────────────────────────────────

/** @private */
function exceedsExcelPrecision(value: unknown): boolean {
  const str     = typeof value === 'string' ? value : String(Number(value));
  const intPart = str.replace(/^-/, '').split('.')[0];
  return intPart.replace(/^0+/, '').length > EXCEL_MAX_SIG_DIGITS;
}

/** @private */
function castCell(value: unknown, type: ColumnType): unknown {
  if (value === null || value === undefined) return null;

  if (type === 'number') {
    const num = Number(value);
    if (isNaN(num) || exceedsExcelPrecision(value)) return String(value);
    return num;
  }
  if (type === 'date' || type === 'datetime') {
    const d = value instanceof Date ? value : new Date(value as string);
    return isNaN(d.getTime()) ? String(value) : d;
  }
  return String(value);
}

/** @private */
function buildColumnSpec(colDefs: ColumnDef[]): Partial<ExcelJS.Column>[] {
  return colDefs.map((col) => {
    const style: Partial<ExcelJS.Style> = {};
    const numFmt = col.numFmt ?? (
      col.type === 'date'     ? 'dd/mm/yyyy' :
      col.type === 'datetime' ? 'dd/mm/yyyy hh:mm:ss' :
      null
    );

    if (numFmt) style.numFmt = numFmt;

    if (col.align || col.wrapText) {
      style.alignment = {};
      if (col.align)    style.alignment.horizontal = col.align;
      if (col.wrapText) style.alignment.wrapText   = true;
    }

    if (col.bgColor) {
      style.fill = {
        type   : 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF' + col.bgColor.replace(/^#/, '') },
      } as ExcelJS.Fill;
    }

    if (col.fontColor) {
      style.font = { color: { argb: 'FF' + col.fontColor.replace(/^#/, '') } };
    }

    const spec: Partial<ExcelJS.Column> = { key: col.key, width: col.width ?? 18 };
    if (Object.keys(style).length) spec.style = style as ExcelJS.Style;
    return spec;
  });
}

/** @private */
function writeHeaderRow(
  ws          : StreamWorksheet,
  colDefs     : ColumnDef[],
  headerStyle : HeaderStyle | null = null
): void {
  const row = ws.addRow(colDefs.map((col) => col.header ?? col.key));

  row.font = { bold: headerStyle?.bold ?? true };
  if (headerStyle?.fontColor) {
    row.font = { ...row.font, color: { argb: 'FF' + headerStyle.fontColor.replace(/^#/, '') } };
  }
  if (headerStyle?.bgColor) {
    row.fill = {
      type   : 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF' + headerStyle.bgColor.replace(/^#/, '') },
    } as ExcelJS.Fill;
  }

  row.commit();
}

/**
 * Map Oracle dbTypeName to OracleSqlToExcel column type.
 * @private
 * @param dbTypeName - e.g. 'NUMBER', 'VARCHAR2', 'DATE', 'TIMESTAMP'
 */
function oracleTypeToColType(dbTypeName: string | undefined): ColumnType {
  if (!dbTypeName) return 'text';
  const t = dbTypeName.toUpperCase();
  if (t === 'NUMBER' || t === 'BINARY_DOUBLE' || t === 'BINARY_FLOAT') return 'number';
  if (t === 'DATE')                                                      return 'date';
  if (t.startsWith('TIMESTAMP'))                                         return 'datetime';
  return 'text';
}

/**
 * Apply DocHeaderStyle to a single ExcelJS cell.
 * @private
 */
function applyDocHeaderCellStyle(cell: ExcelJS.Cell, style: DocHeaderStyle | undefined): void {
  if (!style) return;

  const font: Partial<ExcelJS.Font> = {
    ...(style.bold      && { bold: true }),
    ...(style.italic    && { italic: true }),
    ...(style.fontSize  && { size: style.fontSize }),
    ...(style.fontColor && { color: { argb: 'FF' + style.fontColor.replace(/^#/, '') } }),
  };

  if (Object.keys(font).length) cell.font = font as ExcelJS.Font;

  if (style.align) cell.alignment = { horizontal: style.align, vertical: 'middle' };

  if (style.bgColor) {
    cell.fill = {
      type   : 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF' + style.bgColor.replace(/^#/, '') },
    } as ExcelJS.Fill;
  }
}

/**
 * Write all doc-header rows into a streaming worksheet.
 * Supports simple mode (text + full-width merge) and column mode (DocHeaderCell[]).
 * All mergeCells calls are made before row.commit() as required by ExcelJS streaming.
 * @private
 * @param ws            - Streaming worksheet writer.
 * @param docHeaderRows - Rows to write.
 * @param colCount      - Number of data columns (used for full-width merge in simple mode).
 * @returns Number of rows written.
 */
function writeDocHeaderRows(
  ws            : StreamWorksheet,
  docHeaderRows : DocHeaderRow[],
  colCount      : number
): number {
  // blockedUntil[col] = last row index still occupied by a mergeDown from a previous row
  const blockedUntil: Record<number, number> = {};

  for (const docHeaderRow of docHeaderRows) {
    const row    = ws.addRow([]);
    const rowNum = row.number;

    if (docHeaderRow.height) row.height = docHeaderRow.height;

    if (Array.isArray(docHeaderRow.columns)) {
      // ── Multi-column mode ─────────────────────────────────────────────────────
      let col = 1;
      for (const cellDef of docHeaderRow.columns) {
        // Skip columns blocked by a mergeDown from a previous row
        while (col <= colCount && (blockedUntil[col] ?? 0) >= rowNum) col++;
        if (col > colCount) break;

        const mergeAcross = cellDef.mergeAcross ?? 0;
        const mergeDown   = cellDef.mergeDown   ?? 0;
        const endCol      = col + mergeAcross;
        const endRow      = rowNum + mergeDown;

        row.getCell(col).value = cellDef.text ?? '';
        applyDocHeaderCellStyle(row.getCell(col), cellDef.style);

        if (mergeAcross > 0 || mergeDown > 0) {
          ws.mergeCells(rowNum, col, endRow, endCol);
        }

        // Mark columns as blocked for subsequent rows
        if (mergeDown > 0) {
          for (let c = col; c <= endCol; c++) {
            blockedUntil[c] = endRow;
          }
        }

        col = endCol + 1;
      }
    } else {
      // ── Simple mode (single text, full-width merge or merge=false) ────────────
      row.getCell(1).value = docHeaderRow.text ?? '';
      applyDocHeaderCellStyle(row.getCell(1), docHeaderRow.style);

      if ((docHeaderRow.merge ?? true) && colCount > 1) {
        ws.mergeCells(rowNum, 1, rowNum, colCount);
      }
    }

    row.commit();
  }

  return docHeaderRows.length;
}

/** @private */
function resolveSheetName(sheetName: string | string[], index: number): string {
  if (Array.isArray(sheetName)) {
    if (index < sheetName.length) return sheetName[index];
    const base   = sheetName[sheetName.length - 1];
    const suffix = index - sheetName.length + 2;
    return `${base} ${suffix}`;
  }
  return index === 0 ? sheetName : `${sheetName} ${index + 1}`;
}

// ── SheetConfig ───────────────────────────────────────────────────────────────

/**
 * Per-sheet configuration used inside a `.sheet()` callback.
 * Do not instantiate directly — always obtain via {@link OracleSqlToExcelBuilder#sheet}.
 *
 * @example
 * OracleSqlToExcel()
 *   .sheet('Detail', s => s
 *     .sql('SELECT * FROM MY_TABLE')
 *     .columns([{ key: 'BALANCE', type: 'number', numFmt: '#,##0.00' }])
 *     .freezeHeader()
 *     .autoFilter()
 *     .headerStyle({ bgColor: '4472C4', fontColor: 'FFFFFF' })
 *     .onRowError('skip')
 *   )
 *   .run();
 */
class SheetConfig {
  /** @private */ _name             : string | string[];
  /** @private */ _sql              : string;
  /** @private */ _param            : Record<string, unknown>;
  /** @private */ _executeOptions   : Record<string, unknown>;
  /** @private */ _columns          : ColumnDef[];
  /** @private */ _maxRowsPerSheet  : number;
  /** @private */ _fetchSize        : number;
  /** @private */ _freezeHeader     : boolean;
  /** @private */ _autoFilter       : boolean;
  /** @private */ _headerStyle      : HeaderStyle | null;
  /** @private */ _onRowError       : 'throw' | 'skip';
  /** @private */ _docHeader        : DocHeaderRow[];
  /** @private */ _showTotalRows    : boolean;
  /** @private */ _resolvedTotalRows: number | null;

  constructor(name: string | string[]) {
    this._name              = name;
    this._sql               = '';
    this._param             = {};
    this._executeOptions    = {};
    this._columns           = [];
    this._maxRowsPerSheet   = EXCEL_MAX_ROWS;
    this._fetchSize         = 50000;
    this._freezeHeader      = false;
    this._autoFilter        = false;
    this._headerStyle       = null;
    this._onRowError        = 'throw';
    this._docHeader         = [];
    this._showTotalRows     = false;
    this._resolvedTotalRows = null;
  }

  /**
   * SQL query for this sheet, with optional bind parameters and execute options.
   *
   * @param query          - Oracle SQL query.
   * @param param          - Oracle bind parameters, e.g. `{ id: 1 }`.
   * @param executeOptions - Override Oracle execute options for this sheet only,
   *                         e.g. `{ autoCommit: false }`. Merged over workbook-level
   *                         options; `outFormat`, `resultSet`, and `fetchArraySize`
   *                         are always overridden internally.
   *
   * @example
   * s.sql('SELECT * FROM T WHERE CODE = :code', { code: '019' })
   * s.sql('SELECT * FROM T', {}, { autoCommit: false })
   */
  sql(
    query          : string,
    param          : Record<string, unknown> = {},
    executeOptions : Record<string, unknown> = {}
  ): this {
    this._sql            = query;
    this._param          = param;
    this._executeOptions = executeOptions;
    return this;
  }

  /**
   * Column definitions. When omitted, all DB columns are written using their Oracle key names as headers.
   * @param value - OPTIONAL. Default: all DB columns auto-detected from Oracle metadata.
   */
  columns(value: ColumnDef[]): this { this._columns = value; return this; }

  /**
   * Maximum data rows before a new sheet is created.
   * When exceeded, a new physical sheet is added automatically with an incremented name.
   * @param value - OPTIONAL. Default: `1_000_000` (Excel row limit).
   */
  maxRowsPerSheet(value: number): this { this._maxRowsPerSheet = value; return this; }

  /**
   * Number of rows fetched from Oracle per round-trip.
   * Increase for large exports to reduce round-trips; decrease to lower memory usage.
   * @param value - OPTIONAL. Default: `50_000`.
   */
  fetchSize(value: number): this { this._fetchSize = value; return this; }

  /**
   * Freeze the header row so it stays visible when scrolling.
   * @param value - OPTIONAL. Default: `true` when called without argument.
   */
  freezeHeader(value = true): this { this._freezeHeader = value; return this; }

  /**
   * Add an Excel auto-filter dropdown to every column in the header row.
   * @param value - OPTIONAL. Default: `true` when called without argument.
   */
  autoFilter(value = true): this { this._autoFilter = value; return this; }

  /**
   * Override the default header row style. Default style: bold text, no background color.
   * @param value - OPTIONAL. Omit to keep default bold style.
   *
   * @example
   * .headerStyle({ bgColor: '4472C4', fontColor: 'FFFFFF' }) // blue bg, white text
   */
  headerStyle(value: HeaderStyle): this { this._headerStyle = value; return this; }

  /**
   * Behaviour when a single row fails to cast or write.
   * @param value - `'throw'` (default): aborts on first row error.
   *               `'skip'`: drops the bad row; count reported in `result.skippedRows`.
   */
  onRowError(value: 'throw' | 'skip'): this { this._onRowError = value; return this; }

  /**
   * Document header — one or more rows written **before** the table header row,
   * **on the first sheet only** (not on continuation/split sheets).
   *
   * Split sheets automatically prepend a *"Continued from sheet: <name>"* row instead.
   *
   * @param rows - Array of doc-header rows. Each element follows the {@link DocHeaderRow} interface.
   *
   * @example
   * // Simple mode
   * s.docHeader([
   *   { text: 'ACME CORPORATION', style: { bold: true, fontSize: 14, align: 'center' }, height: 24 },
   *   { text: 'Period: January 2026', style: { italic: true, align: 'center' } },
   *   { text: '' },
   * ])
   *
   * @example
   * // Column mode — multi-cell per row with horizontal & vertical merge
   * s.docHeader([
   *   {
   *     columns: [
   *       { text: 'LOGO', mergeAcross: 1, mergeDown: 1, style: { align: 'center', bold: true } },
   *       { text: 'BANK STATEMENT REPORT', mergeAcross: 8, style: { bold: true, fontSize: 14, align: 'center' } },
   *     ],
   *     height: 30,
   *   },
   * ])
   */
  docHeader(rows: DocHeaderRow[]): this { this._docHeader = rows; return this; }

  /**
   * Prepend a row count summary above the table header on every sheet (including splits).
   *
   * Format: `"Showing rows X – Y of Z total"`
   *
   * The library automatically runs `SELECT COUNT(*) FROM (<sql>)` in parallel on a separate
   * connection before streaming begins. There is no overhead during the stream itself.
   *
   * Silently skipped when:
   * - The SQL uses a CTE (`WITH ... AS`) — cannot be wrapped in a subquery.
   * - The COUNT query fails — export continues normally without the summary row.
   *
   * @param value - OPTIONAL. Default when called without argument: `true`.
   */
  showTotalRows(value = true): this { this._showTotalRows = value; return this; }
}

// ── FileConfig ────────────────────────────────────────────────────────────────

/**
 * Per-file configuration used inside a `.file()` callback.
 * Groups one or more sheets into a single logical output file.
 * When `.maxRowsPerFile()` is set, the file is automatically split into
 * multiple physical files when the row limit is reached.
 *
 * Do not instantiate directly — always obtain via {@link OracleSqlToExcelBuilder#file}.
 *
 * @example
 * OracleSqlToExcel()
 *   .file('report', f => f
 *     .maxRowsPerFile(1_000_000)
 *     .sheet('Detail',  s => s.sql(SQL1).columns(COLS1).maxRowsPerSheet(900_000))
 *     .sheet('Summary', s => s.sql(SQL2).columns(COLS2))
 *   )
 *   .run()
 */
class FileConfig {
  /** @private */ _name          : string;
  /** @private */ _maxRowsPerFile: number;
  /** @private */ _sheets        : SheetConfig[];

  constructor(name: string) {
    this._name           = name;
    this._maxRowsPerFile = 0;
    this._sheets         = [];
  }

  /**
   * Split this file into multiple physical `.xlsx` files when data rows exceed `value`.
   * Each file is named `<name>_<startRow>-<endRow>.xlsx` (single sheet) or
   * `<name>_1.xlsx`, `<name>_2.xlsx`, … (multiple sheets).
   * @param value - OPTIONAL. Default: `0` (no split — single file).
   */
  maxRowsPerFile(value: number): this { this._maxRowsPerFile = value; return this; }

  /**
   * Add a sheet to this file. Identical API to {@link OracleSqlToExcelBuilder#sheet}.
   */
  sheet(name: string | string[], fn: (s: SheetConfig) => void): this {
    const names = Array.isArray(name) ? name : [name];
    for (const n of names) {
      if (n.length > 25) throw new Error(`Sheet name "${n}" exceeds 25 characters (${n.length}). Shorten the name.`);
    }
    const cfg = new SheetConfig(name);
    fn(cfg);
    this._sheets.push(cfg);
    return this;
  }
}

// ── Builder ───────────────────────────────────────────────────────────────────

/**
 * Fluent builder that streams Oracle SQL results directly into an Excel workbook.
 * Configure each sheet via `.sheet(name, fn)`, then call a terminal method.
 *
 * Terminal methods: `.run()` → file on disk · `.pipe(stream)` → any writable · `.toBuffer()` → `Buffer`.
 *
 * @example
 * // Single sheet — write to file
 * const { file } = await OracleSqlToExcel()
 *   .sheet('Report', s => s
 *     .sql('SELECT * FROM MY_TABLE')
 *     .columns([{ key: 'BALANCE', type: 'number', numFmt: '#,##0.00' }])
 *     .freezeHeader()
 *     .headerStyle({ bgColor: '4472C4', fontColor: 'FFFFFF' })
 *   )
 *   .filePrefix('my-report')
 *   .run();
 *
 * @example
 * // Multi-sheet with progress
 * const { file } = await OracleSqlToExcel()
 *   .sheet('Summary', s => s.sql(SQL1).columns(COLS1).freezeHeader().autoFilter())
 *   .sheet('Detail',  s => s.sql(SQL2).columns(COLS2).onRowError('skip'))
 *   .filePrefix('report')
 *   .onProgress(({ sheet, rowsWritten }) => console.log(sheet, rowsWritten))
 *   .run();
 *
 * @example
 * // Stream to Express response
 * res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
 * res.setHeader('Content-Disposition', 'attachment; filename="report.xlsx"');
 * await OracleSqlToExcel()
 *   .sheet('Data', s => s.sql(SQL).columns(COLS))
 *   .pipe(res);
 */
class OracleSqlToExcelBuilder {
  /** @private */ _connectionFactory      : (() => Promise<OracleConnection>) | null;
  /** @private */ _executeOptions         : Record<string, unknown>;
  /** @private */ _outputDir              : string;
  /** @private */ _filePrefix             : string;
  /** @private */ _compress               : boolean;
  /** @private */ _debug                  : boolean;
  /** @private */ _onProgressCb           : ((info: ProgressInfo) => void) | null;
  /** @private */ _sheets                 : SheetConfig[];
  /** @private */ _files                  : FileConfig[];
  /** @private */ _backpressureThreshold  : number;
  /** @private */ _asZip                  : boolean;

  constructor() {
    this._connectionFactory     = null;
    this._executeOptions        = {};
    this._outputDir             = process.cwd();
    this._filePrefix            = 'export';
    this._compress              = false;
    this._debug                 = false;
    this._onProgressCb          = null;
    this._sheets                = [];
    this._files                 = [];
    this._backpressureThreshold = 512 * 1024 * 1024; // 512 MB
    this._asZip                 = false;
  }

  // ── Workbook-level methods ─────────────────────────────────────────────────

  /**
   * Factory function called by the library each time a new connection is needed
   * (COUNT queries and the main stream). Must return a Promise that resolves to an
   * Oracle connection. The library automatically closes every connection when done.
   *
   * @example
   * .connectionFactory(() => oracledb.getPool('myPool').getConnection())
   */
  connectionFactory(fn: () => Promise<OracleConnection>): this {
    this._connectionFactory = fn;
    return this;
  }

  /**
   * Default Oracle `connection.execute` options applied to **all sheets**.
   * Per-sheet options set via the third argument of `.sql()` are merged on top of these.
   * `outFormat`, `resultSet`, and `fetchArraySize` are always overridden internally.
   * @param value - OPTIONAL. Default: `{}`.
   */
  executeOptions(value: Record<string, unknown>): this { this._executeOptions = value; return this; }

  /**
   * Destination directory for `.run()`.
   * @param value - OPTIONAL. Default: `process.cwd()`.
   */
  outputDir(value: string): this { this._outputDir = value; return this; }

  /**
   * Output filename **without** extension for `.run()`. The file will be saved as `<value>.xlsx`.
   * @param value - OPTIONAL. Default: `'export'`.
   */
  filePrefix(value: string): this { this._filePrefix = value; return this; }

  /**
   * Enable or disable ZIP compression inside the XLSX file.
   *
   * Default is `false` (no compression) because:
   * - Compression is CPU-intensive and blocks the stream flush, causing slow download speed.
   * - For HTTP streaming, the response itself can be gzip-compressed by the web server.
   * - For `.run()` (file on disk), enable compression to reduce file size at the cost of speed.
   *
   * @param value - OPTIONAL. Default when called without argument: `true`.
   *
   * @example
   * .compress(false)  // fastest — recommended for .pipe(res) over HTTP
   * .compress(true)   // smaller file — recommended for .run() to disk
   */
  compress(value = true): this { this._compress = value; return this; }

  /**
   * Maximum process RSS (Resident Set Size) allowed during `.pipe()` streaming before the
   * Oracle fetch is paused. When `process.memoryUsage().rss` exceeds this value after a batch,
   * the library polls every 200 ms and waits until RSS drops below the threshold before
   * fetching the next Oracle batch.
   *
   * This guards against memory exhaustion when the output stream is behind a reverse proxy
   * (e.g. nginx) — where Node.js `write()` never returns `false` even though the end client
   * is downloading slowly and data is accumulating in the Node.js output buffer.
   *
   * Has no effect for `.run()` (file) or `.toBuffer()`.
   * @param bytes - Default: `536870912` (512 MB).
   *
   * @example
   * .backpressureThreshold(256 * 1024 * 1024)  // pause when RSS exceeds 256 MB
   */
  backpressureThreshold(bytes: number): this { this._backpressureThreshold = bytes; return this; }

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

  /**
   * Add a logical file to the export. Each `.file()` call defines one output `.xlsx` file
   * (or a set of split files if `.maxRowsPerFile()` is set inside the callback).
   *
   * Only applies to `.run()`. Using `.file()` with `.pipe()` or `.toBuffer()` throws an error.
   *
   * @param name - Base filename without extension. Combined with `.outputDir()` to form the path.
   * @param fn   - Callback that configures the file via a {@link FileConfig} instance.
   *
   * @example
   * OracleSqlToExcel()
   *   .connectionFactory(() => pool.getConnection())
   *   .outputDir('/tmp')
   *   .file('laporan', f => f
   *     .maxRowsPerFile(1_000_000)
   *     .sheet('Detail',  s => s.sql(SQL1).columns(COLS1).maxRowsPerSheet(900_000))
   *     .sheet('Summary', s => s.sql(SQL2).columns(COLS2))
   *   )
   *   .run()
   */
  file(name: string, fn: (f: FileConfig) => void): this {
    const cfg = new FileConfig(name);
    fn(cfg);
    this._files.push(cfg);
    return this;
  }

  /**
   * Enable verbose debug logging to `console.log` at each execution stage.
   * Logs are emitted only when **both** conditions are met:
   *   1. `.debug(true)` has been called.
   *   2. `NODE_ENV` is **not** `production` or `prod`.
   * Errors are always logged to `console.error` regardless of this setting.
   * @param value - OPTIONAL. Default when called without argument: `true`.
   */
  debug(value = true): this { this._debug = value; return this; }

  /**
   * Callback invoked after each fetch batch is written.
   * Useful for WebSocket / SSE progress updates on large exports.
   *
   * @example
   * .onProgress(({ sheet, rowsWritten, skippedRows, totalRowsWritten }) => {
   *   io.emit('progress', { sheet, rowsWritten, skippedRows, totalRowsWritten });
   * })
   */
  onProgress(cb: (info: ProgressInfo) => void): this { this._onProgressCb = cb; return this; }

  /**
   * Add a sheet to the workbook. Each call appends one logical sheet (which may
   * split into multiple physical sheets when `maxRowsPerSheet` is exceeded).
   *
   * @param name - Sheet name. Pass an array to control names on splits:
   *   `['Part A', 'Part B']` → 'Part A', 'Part B', 'Part B 2', …
   * @param fn   - Callback that receives and configures the sheet.
   *
   * @example
   * OracleSqlToExcel()
   *   .sheet('Summary', s => s.sql(SQL1).columns(COLS1).freezeHeader().autoFilter())
   *   .sheet('Detail',  s => s.sql(SQL2).maxRowsPerSheet(500_000).onRowError('skip'))
   *   .run();
   */
  sheet(name: string | string[], fn: (s: SheetConfig) => void): this {
    const names = Array.isArray(name) ? name : [name];
    for (const n of names) {
      if (n.length > 25) {
        throw new Error(`Sheet name "${n}" exceeds 25 characters (${n.length}). Shorten the name.`);
      }
    }
    const cfg = new SheetConfig(name);
    fn(cfg);
    this._sheets.push(cfg);
    return this;
  }

  // ── Internal ───────────────────────────────────────────────────────────────

  /**
   * Streams one sheet's SQL result into the workbook.
   * @private
   */
  async _executeSheet(
    connection  : OracleConnection,
    workbook    : StreamWorkbook,
    sheetCfg    : SheetConfig,
    progressCtx : ProgressCtx,
    drainFn?    : (() => Promise<void>) | null
  ): Promise<{ sheetNames: string[]; skippedRows: number }> {
    const sheetNames                            : string[]          = [];
    let   sheetIndex                                                = 0;
    let   rowCounter                                                = 0;
    let   totalRows                                                 = 0;
    let   skippedRows                                               = 0;
    let   resolvedColDefs : ColumnDef[] | null                      = sheetCfg._columns.length > 0 ? sheetCfg._columns : null;
    // worksheet is assigned by createNewSheet() before first use — definite assignment
    let   worksheet       : StreamWorksheet                         = null!;
    let   resultSet       : OracleResultSet | null                  = null;

    const isDevEnv = !['production', 'prod'].includes((process.env.NODE_ENV ?? '').toLowerCase());
    const dbg      = (msg: string): void => { if (this._debug && isDevEnv) console.log(`[OracleSqlToExcel:DEBUG] ${msg}`); };

    const createNewSheet = (): void => {
      const name = resolveSheetName(sheetCfg._name, sheetIndex);
      dbg(`createNewSheet — sheetIndex=${sheetIndex} name="${name}"`);

      if (name.length > 25) {
        throw new Error(`Sheet name "${name}" exceeds 25 characters (${name.length}). Shorten the name or reduce maxRowsPerSheet.`);
      }

      worksheet = workbook.addWorksheet(name);
      sheetNames.push(name);
      dbg(`  addWorksheet OK — "${name}"`);

      const colCount      = resolvedColDefs ? resolvedColDefs.length : 1;
      let   prependedRows = 0;
      dbg(`  colCount=${colCount} resolvedColDefs=${resolvedColDefs ? resolvedColDefs.length + ' cols' : 'null'}`);

      if (resolvedColDefs) {
        worksheet.columns = buildColumnSpec(resolvedColDefs);
        dbg('  worksheet.columns set');
      }

      // Doc header: first sheet only (not on split continuations)
      if (sheetIndex === 0 && sheetCfg._docHeader?.length > 0) {
        dbg(`  docHeader: ${sheetCfg._docHeader.length} row(s)`);
        prependedRows = writeDocHeaderRows(worksheet, sheetCfg._docHeader, colCount);
        dbg(`  docHeader done — prependedRows=${prependedRows}`);
      }

      // Continuation notice: split sheets (sheetIndex > 0)
      if (sheetIndex > 0) {
        const prevName = resolveSheetName(sheetCfg._name, sheetIndex - 1);
        dbg(`  prevInfo: "Continued from sheet: ${prevName}"`);
        const prevRow = worksheet.addRow([`Continued from sheet: ${prevName}`]);
        prevRow.font  = { italic: true, color: { argb: 'FF808080' } };
        if (colCount > 1) worksheet.mergeCells(prevRow.number, 1, prevRow.number, colCount);
        prevRow.commit();
        prependedRows++;
        dbg('  prevInfo committed');
      }

      // Row range summary — all sheets including splits
      if (sheetCfg._resolvedTotalRows != null) {
        const fmt    = (n: number): string => n.toLocaleString('en-US');
        const start  = sheetIndex * sheetCfg._maxRowsPerSheet + 1;
        const end    = Math.min((sheetIndex + 1) * sheetCfg._maxRowsPerSheet, sheetCfg._resolvedTotalRows);
        const text   = `Showing rows ${fmt(start)} – ${fmt(end)} of ${fmt(sheetCfg._resolvedTotalRows)} total`;
        dbg(`  rangeInfo: "${text}"`);
        const rangeRow = worksheet.addRow([text]);
        rangeRow.font  = { italic: true, color: { argb: 'FF404040' } };
        if (colCount > 1) worksheet.mergeCells(rangeRow.number, 1, rangeRow.number, colCount);
        rangeRow.commit();
        prependedRows++;
      }

      if (resolvedColDefs) {
        const headerRowNum = prependedRows + 1;
        dbg(`  headerRowNum=${headerRowNum}`);

        if (sheetCfg._freezeHeader) {
          worksheet.views = [{ state: 'frozen', ySplit: headerRowNum }];
          dbg(`  freezeHeader ySplit=${headerRowNum}`);
        }
        if (sheetCfg._autoFilter) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (worksheet as any).autoFilter = {
            from: { row: headerRowNum, column: 1 },
            to  : { row: headerRowNum, column: resolvedColDefs.length },
          };
          dbg('  autoFilter set');
        }

        writeHeaderRow(worksheet, resolvedColDefs, sheetCfg._headerStyle);
        dbg('  headerRow committed');
      }

      dbg(`createNewSheet done — "${name}"`);
    };

    try {
      const sheetLabel = Array.isArray(sheetCfg._name) ? sheetCfg._name[0] : sheetCfg._name;
      dbg(`_executeSheet START — sheet="${sheetLabel}" sql="${sheetCfg._sql?.trim().slice(0, 80)}..."`);

      // When columns are defined, wrap the SQL as a subquery and SELECT only those columns.
      // Strip trailing semicolons so the query is valid as an Oracle inline view.
      // Column names are unquoted — safe for standard uppercase Oracle identifiers.
      // CTE (WITH ...) cannot be wrapped in an inline view — fall back to the original SQL.
      let sql = sheetCfg._sql;
      if (resolvedColDefs) {
        const trimmed = sheetCfg._sql.trim().replace(/;+$/, '');
        const hasCte  = /^\s*WITH\s+/i.test(trimmed);

        let colList: string | null = null;
        if (!hasCte) {
          colList = resolvedColDefs.map((c) => c.key).join(', ');
          sql     = `SELECT ${colList} FROM (${trimmed})`;
        }
        dbg(`SQL wrap — hasCte=${hasCte} cols=${colList ?? '(CTE fallback)'}`);
      }

      dbg('connection.execute ...');
      const execResult = await connection.execute(sql, sheetCfg._param, {
        autoCommit    : true,
        ...this._executeOptions,       // workbook-level defaults
        ...sheetCfg._executeOptions,   // sheet-level overrides (from 3rd arg of .sql())
        outFormat     : OUT_FORMAT_OBJECT,
        resultSet     : true,
        fetchArraySize: sheetCfg._fetchSize,
      });
      dbg(`connection.execute OK — metaData cols=${execResult.metaData?.length ?? 0}`);

      resultSet = execResult.resultSet ?? null;

      // Use Oracle metadata to resolve column names upfront — works even for empty result sets
      if (!resolvedColDefs && execResult.metaData?.length) {
        resolvedColDefs = execResult.metaData.map((m) => ({
          key   : m.name,
          header: m.name,
          type  : oracleTypeToColType(m.dbTypeName),
        }));
        dbg(`auto-detect columns: ${resolvedColDefs.map((c) => c.key).join(', ')}`);

        const isProd = ['production', 'prod'].includes((process.env.NODE_ENV ?? '').toLowerCase());
        if (!isProd) {
          console.warn(
            `[OracleSqlToExcel] WARNING — Sheet "${sheetLabel}": no explicit column definitions set via .columns(). ` +
            `Columns auto-detected from Oracle metadata (${resolvedColDefs.length} column(s)). ` +
            'Column order depends on the SELECT clause and may change if the query or schema changes. ' +
            'Define columns explicitly with .columns([...]) for stable, long-term output.',
          );
        }
      }

      createNewSheet();

      dbg('fetching first batch...');
      let rows = await resultSet!.getRows(sheetCfg._fetchSize);
      dbg(`getRows → ${rows.length} row(s)`);

      while (rows.length > 0) {
        const batchRowsBefore = totalRows;

        for (const row of rows) {
          let rowWritten = false;
          try {
            const rowData: Record<string, unknown> = {};
            resolvedColDefs!.forEach(({ key, type = 'text' }) => {
              rowData[key] = castCell(row[key], type);
            });
            worksheet.addRow(rowData).commit();
            rowWritten = true;
          } catch (rowErr) {
            const msg = rowErr instanceof Error ? rowErr.message : String(rowErr);
            dbg(`rowErr: ${msg} — onRowError=${sheetCfg._onRowError}`);
            if (sheetCfg._onRowError === 'throw') throw rowErr;
            skippedRows++;
          }

          if (rowWritten) {
            rowCounter++;
            totalRows++;

            if (rowCounter >= sheetCfg._maxRowsPerSheet) {
              const nextSheetName = resolveSheetName(sheetCfg._name, sheetIndex + 1);
              const colCount      = resolvedColDefs ? resolvedColDefs.length : 1;
              dbg(`split sheet — rowCounter=${rowCounter} maxRowsPerSheet=${sheetCfg._maxRowsPerSheet} nextSheet="${nextSheetName}"`);

              worksheet.addRow(['']).commit();
              const infoRow = worksheet.addRow([`Continued on sheet: ${nextSheetName}`]);
              infoRow.font  = { italic: true, color: { argb: 'FF404040' } };
              if (colCount > 1) worksheet.mergeCells(infoRow.number, 1, infoRow.number, colCount);
              infoRow.commit();

              dbg(`await worksheet.commit() — "${resolveSheetName(sheetCfg._name, sheetIndex)}"`);
              await worksheet.commit();
              dbg('worksheet.commit OK');

              sheetIndex++;
              rowCounter = 0;
              createNewSheet();
            }
          }
        }

        if (this._onProgressCb) {
          progressCtx.totalRowsWritten += totalRows - batchRowsBefore;
          const label = Array.isArray(sheetCfg._name) ? sheetCfg._name[0] : sheetCfg._name;
          this._onProgressCb({ sheet: label, rowsWritten: totalRows, skippedRows, totalRowsWritten: progressCtx.totalRowsWritten });
        }

        if (drainFn) await drainFn();

        rows = await resultSet!.getRows(sheetCfg._fetchSize);
        dbg(`getRows → ${rows.length} row(s) (totalRows=${totalRows})`);
      }

      dbg(`_executeSheet DONE — sheet="${sheetLabel}" totalRows=${totalRows} skippedRows=${skippedRows}`);
      await worksheet.commit();
      return { sheetNames, skippedRows };
    } catch (err) {
      const msg   = err instanceof Error ? err.message : String(err);
      const stack = err instanceof Error ? err.stack   : undefined;
      console.error(`[OracleSqlToExcel:ERROR] _executeSheet FAILED — ${msg}`);
      if (stack) console.error(stack);
      throw err;
    } finally {
      if (resultSet) await resultSet.close().catch(() => {});
    }
  }

  /**
   * Like _executeSheet but stops after maxRows data rows, returning any unprocessed
   * rows and the open ResultSet so the caller can continue in the next file.
   * @private
   */
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
    const sheetNames                          : string[]        = [];
    let   sheetIndex                                            = 0;
    let   rowCounter                                            = 0;
    let   totalRows                                             = 0;
    let   skippedRows                                           = 0;
    let   fileRowsWritten                                       = 0;
    let   resolvedColDefs : ColumnDef[] | null                  = sheetCfg._columns.length > 0 ? sheetCfg._columns : null;
    let   worksheet       : StreamWorksheet                     = null!;
    let   resultSet       : OracleResultSet | null              = null;
    let   earlyReturn                                           = false;
    let   overflowRows    : Record<string, unknown>[]           = [];

    const createNewSheet = (): void => {
      const name = resolveSheetName(sheetCfg._name, sheetIndex);
      if (name.length > 25) throw new Error(`Sheet name "${name}" exceeds 25 characters (${name.length}).`);
      worksheet = workbook.addWorksheet(name);
      sheetNames.push(name);
      const colCount = resolvedColDefs ? resolvedColDefs.length : 1;
      let prependedRows = 0;
      if (resolvedColDefs) worksheet.columns = buildColumnSpec(resolvedColDefs);
      if (sheetIndex === 0 && sheetCfg._docHeader?.length > 0) {
        prependedRows = writeDocHeaderRows(worksheet, sheetCfg._docHeader, colCount);
      }
      if (sheetIndex > 0) {
        const prevName = resolveSheetName(sheetCfg._name, sheetIndex - 1);
        const prevRow  = worksheet.addRow([`Continued from sheet: ${prevName}`]);
        prevRow.font   = { italic: true, color: { argb: 'FF808080' } };
        if (colCount > 1) worksheet.mergeCells(prevRow.number, 1, prevRow.number, colCount);
        prevRow.commit();
        prependedRows++;
      }
      if (sheetCfg._resolvedTotalRows != null) {
        const fmt      = (n: number): string => n.toLocaleString('en-US');
        const start    = globalRowOffset + sheetIndex * sheetCfg._maxRowsPerSheet + 1;
        const end      = Math.min(globalRowOffset + (sheetIndex + 1) * sheetCfg._maxRowsPerSheet, sheetCfg._resolvedTotalRows);
        const rangeRow = worksheet.addRow([`Showing rows ${fmt(start)} – ${fmt(end)} of ${fmt(sheetCfg._resolvedTotalRows)} total`]);
        rangeRow.font  = { italic: true, color: { argb: 'FF404040' } };
        if (colCount > 1) worksheet.mergeCells(rangeRow.number, 1, rangeRow.number, colCount);
        rangeRow.commit();
        prependedRows++;
      }
      if (resolvedColDefs) {
        const headerRowNum = prependedRows + 1;
        if (sheetCfg._freezeHeader) worksheet.views = [{ state: 'frozen', ySplit: headerRowNum }];
        if (sheetCfg._autoFilter) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (worksheet as any).autoFilter = { from: { row: headerRowNum, column: 1 }, to: { row: headerRowNum, column: resolvedColDefs.length } };
        }
        writeHeaderRow(worksheet, resolvedColDefs, sheetCfg._headerStyle);
      }
    };

    try {
      const sheetLabel = Array.isArray(sheetCfg._name) ? sheetCfg._name[0] : sheetCfg._name;

      if (existingRS) {
        resultSet = existingRS;
      } else {
        let sql = sheetCfg._sql;
        if (resolvedColDefs) {
          const trimmed = sheetCfg._sql.trim().replace(/;+$/, '');
          if (!/^\s*WITH\s+/i.test(trimmed)) sql = `SELECT ${resolvedColDefs.map((c) => c.key).join(', ')} FROM (${trimmed})`;
        }
        const execResult = await connection.execute(sql, sheetCfg._param, {
          autoCommit: true, ...this._executeOptions, ...sheetCfg._executeOptions,
          outFormat: OUT_FORMAT_OBJECT, resultSet: true, fetchArraySize: sheetCfg._fetchSize,
        });
        resultSet = execResult.resultSet ?? null;
        if (!resolvedColDefs && execResult.metaData?.length) {
          resolvedColDefs = execResult.metaData.map((m) => ({ key: m.name, header: m.name, type: oracleTypeToColType(m.dbTypeName) }));
        }
      }

      createNewSheet();

      let rows = pendingRows.length > 0 ? pendingRows : await resultSet!.getRows(sheetCfg._fetchSize);

      while (rows.length > 0) {
        const batchRowsBefore = totalRows;

        for (let i = 0; i < rows.length; i++) {
          if (fileRowsWritten >= maxRows) {
            earlyReturn  = true;
            overflowRows = rows.slice(i);
            break;
          }

          const row = rows[i];
          let rowWritten = false;
          try {
            const rowData: Record<string, unknown> = {};
            resolvedColDefs!.forEach(({ key, type = 'text' }) => { rowData[key] = castCell(row[key], type); });
            worksheet.addRow(rowData).commit();
            rowWritten = true;
          } catch (rowErr) {
            if (sheetCfg._onRowError === 'throw') throw rowErr;
            skippedRows++;
          }

          if (rowWritten) {
            rowCounter++;
            totalRows++;
            fileRowsWritten++;

            if (rowCounter >= sheetCfg._maxRowsPerSheet) {
              const nextSheetName = resolveSheetName(sheetCfg._name, sheetIndex + 1);
              const colCount      = resolvedColDefs ? resolvedColDefs.length : 1;
              worksheet.addRow(['']).commit();
              const infoRow = worksheet.addRow([`Continued on sheet: ${nextSheetName}`]);
              infoRow.font  = { italic: true, color: { argb: 'FF404040' } };
              if (colCount > 1) worksheet.mergeCells(infoRow.number, 1, infoRow.number, colCount);
              infoRow.commit();
              await worksheet.commit();
              sheetIndex++;
              rowCounter = 0;
              createNewSheet();
            }
          }
        }

        if (this._onProgressCb) {
          progressCtx.totalRowsWritten += totalRows - batchRowsBefore;
          this._onProgressCb({ sheet: sheetLabel, rowsWritten: totalRows, skippedRows, totalRowsWritten: progressCtx.totalRowsWritten });
        }

        if (earlyReturn) break;
        if (drainFn) await drainFn();
        rows = await resultSet!.getRows(sheetCfg._fetchSize);
      }

      await worksheet.commit();
      return { sheetNames, skippedRows, rowsWritten: fileRowsWritten, overflowRows, openRS: earlyReturn ? resultSet! : null };
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error(`[OracleSqlToExcel:ERROR] _executeSheetSegment FAILED — ${msg}`);
      throw err;
    } finally {
      if (resultSet && !earlyReturn) await resultSet.close().catch(() => {});
    }
  }

  /**
   * Multi-file execution: splits one sheet's result across multiple .xlsx files.
   * @private
   */
  async _executeFileConfig(cfg: FileConfig): Promise<MultiRunResult> {
    const allSheets   : string[]      = [];
    const allFiles    : FileSegment[] = [];
    const progressCtx : ProgressCtx   = { totalRowsWritten: 0 };
    let   totalSkipped                = 0;
    let   connection  : OracleConnection | null = null;

    // Per-sheet state carried across files
    interface SheetState {
      openRS  : OracleResultSet | null;
      pending : Record<string, unknown>[];
      done    : boolean;
      // row counters for single-sheet filename suffix
      globalStart: number;
      globalEnd  : number;
    }

    try {
      connection = await this._connectionFactory!();

      // COUNT queries
      await Promise.all(cfg._sheets.map(async (sheetCfg) => {
        const label = Array.isArray(sheetCfg._name) ? sheetCfg._name[0] : sheetCfg._name;
        if (!sheetCfg._showTotalRows) return;
        try {
          const trimmed = sheetCfg._sql.trim().replace(/;+$/, '');
          if (/^\s*WITH\s+/i.test(trimmed)) return;
          const countConn = await this._connectionFactory!();
          try {
            const result = await countConn.execute(`SELECT COUNT(*) AS TOTAL FROM (${trimmed})`, sheetCfg._param, { outFormat: OUT_FORMAT_OBJECT });
            sheetCfg._resolvedTotalRows = (result.rows?.[0]?.['TOTAL'] as number | undefined) ?? null;
          } finally {
            await countConn.close().catch(() => {});
          }
        } catch (err) {
          const msg = err instanceof Error ? err.message : String(err);
          console.error(`[OracleSqlToExcel:ERROR] COUNT query failed for sheet "${label}": ${msg}`);
        }
      }));

      const rssThreshold = this._backpressureThreshold;
      const drainFn: (() => Promise<void>) | null = rssThreshold > 0
        ? async () => {
            if (process.memoryUsage().rss <= rssThreshold) return;
            const started = Date.now();
            while (process.memoryUsage().rss > rssThreshold) {
              if (Date.now() - started > 30_000) break;
              await new Promise<void>(r => setTimeout(r, 200));
            }
          }
        : null;

      // Each file contains ALL sheets. Exhausted sheets are skipped in subsequent files.
      // Single-sheet → row-range filename. Multi-sheet → sequential index filename.
      const isSingleSheet = cfg._sheets.length === 1;
      const maxRows       = cfg._maxRowsPerFile > 0 ? cfg._maxRowsPerFile : Number.MAX_SAFE_INTEGER;
      const states        = new Map<SheetConfig, SheetState>(
        cfg._sheets.map((s) => [s, { openRS: null, pending: [], done: false, globalStart: 1, globalEnd: 0 }])
      );

      let fileIndex = 0;

      while ([...states.values()].some((st) => !st.done)) {
        const tempFile = path.join(this._outputDir, `${cfg._name}__tmp_${fileIndex}.xlsx`);

        const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
          filename        : tempFile,
          useStyles       : true,
          useSharedStrings: false,
          zip             : this._compress ? undefined : { zlib: { level: 0 } },
        } as unknown as ExcelJS.stream.xlsx.WorkbookWriterOptions);

        let fileFirstSheetStart = 0;
        let fileFirstSheetEnd   = 0;

        for (const sheetCfg of cfg._sheets) {
          const st = states.get(sheetCfg)!;
          if (st.done) continue;

          const seg = await this._executeSheetSegment(
            connection, workbook, sheetCfg, progressCtx, drainFn,
            maxRows, st.pending, st.openRS,
            st.globalStart - 1
          );

          st.globalEnd = st.globalStart + seg.rowsWritten - 1;
          if (isSingleSheet) {
            fileFirstSheetStart = st.globalStart;
            fileFirstSheetEnd   = st.globalEnd;
          }

          allSheets.push(...seg.sheetNames);
          totalSkipped += seg.skippedRows;

          st.globalStart = st.globalEnd + 1;
          st.openRS      = seg.openRS;
          st.pending     = seg.overflowRows;
          st.done        = seg.openRS === null && seg.overflowRows.length === 0;
        }

        await workbook.commit();

        const finalFile = isSingleSheet
          ? path.join(this._outputDir, `${cfg._name}_${fileFirstSheetStart}-${fileFirstSheetEnd}.xlsx`)
          : path.join(this._outputDir, `${cfg._name}_${fileIndex + 1}.xlsx`);

        await fs.promises.rename(tempFile, finalFile);
        allFiles.push({ file: finalFile, startRow: fileFirstSheetStart, endRow: fileFirstSheetEnd });

        fileIndex++;
      }

      return { success: true, files: allFiles, sheets: allSheets, skippedRows: totalSkipped };
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error(`[OracleSqlToExcel:ERROR] _executeFileConfig FAILED — ${msg}`);
      return { success: false, error: msg, files: allFiles, sheets: allSheets, skippedRows: totalSkipped };
    } finally {
      if (connection) await connection.close().catch(() => {});
    }
  }

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
      if (archiveError) throw archiveError;
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

  /**
   * Core execution logic shared by `.run()`, `.pipe()`, and `.toBuffer()`.
   * @private
   */
  async _execute(workbookTarget: WorkbookTarget): Promise<Result> {
    const allSheets   : string[]     = [];
    const progressCtx : ProgressCtx  = { totalRowsWritten: 0 };

    let connection   : OracleConnection | null = null;
    let workbook     : StreamWorkbook   | null = null;
    let totalSkipped                           = 0;

    const isDevEnv = !['production', 'prod'].includes((process.env.NODE_ENV ?? '').toLowerCase());
    const dbg      = (msg: string): void => { if (this._debug && isDevEnv) console.log(`[OracleSqlToExcel:DEBUG] ${msg}`); };

    try {
      if (!this._connectionFactory) {
        throw new Error('No connection factory set. Call .connectionFactory(() => pool.getConnection()) before running the export.');
      }

      if (this._sheets.length === 0) {
        throw new Error('No sheets defined. Call .sheet(name, fn) at least once before running the export.');
      }

      for (const sheetCfg of this._sheets) {
        if (!sheetCfg._sql?.trim()) {
          const label = Array.isArray(sheetCfg._name) ? sheetCfg._name[0] : sheetCfg._name;
          throw new Error(`Sheet "${label}": no SQL query set. Call .sql() before running the export.`);
        }
      }

      const sheets = this._sheets;
      dbg(`_execute START — ${sheets.length} sheet(s)`);

      // Auto-count: run SELECT COUNT(*) in parallel for every sheet with showTotalRows=true.
      // Each count query uses its own connection so it never blocks the main streaming connection.
      await Promise.all(sheets.map(async (sheetCfg) => {
        const label = Array.isArray(sheetCfg._name) ? sheetCfg._name[0] : sheetCfg._name;
        if (!sheetCfg._showTotalRows) return;

        try {
          const trimmed = sheetCfg._sql.trim().replace(/;+$/, '');
          if (/^\s*WITH\s+/i.test(trimmed)) {
            dbg(`COUNT skip (CTE) — sheet="${label}"`);
            return;
          }

          dbg(`COUNT query START — sheet="${label}"`);
          const countSql  = `SELECT COUNT(*) AS TOTAL FROM (${trimmed})`;
          const countConn = await this._connectionFactory!();
          try {
            const result = await countConn.execute(countSql, sheetCfg._param, { outFormat: OUT_FORMAT_OBJECT });
            sheetCfg._resolvedTotalRows = (result.rows?.[0]?.['TOTAL'] as number | undefined) ?? null;
            dbg(`COUNT query OK — sheet="${label}" total=${sheetCfg._resolvedTotalRows}`);
          } finally {
            await countConn.close().catch(() => {});
          }
        } catch (err) {
          // Count failed — row-range info is skipped for this sheet; export continues normally
          const msg = err instanceof Error ? err.message : String(err);
          dbg(`COUNT query FAILED — sheet="${label}" err="${msg}"`);
          console.error(`[OracleSqlToExcel:ERROR] COUNT query failed for sheet "${label}": ${msg}`);
        }
      }));

      dbg('opening main stream connection ...');
      connection = await this._connectionFactory!();
      dbg('getConnection OK');

      const targetStream = workbookTarget.stream ?? null;
      let drainFn: (() => Promise<void>) | null = null;

      if (targetStream) {
        let needsDrain    = false;
        let streamAborted = false;

        // Intercept write() to capture the authoritative Node.js backpressure signal.
        // writableLength / writableNeedDrain only reflect the final stream buffer and
        // miss data already queued in ExcelJS's internal archiver. write() returning
        // false is the correct signal that downstream cannot accept more data.
        const origWrite = targetStream.write.bind(targetStream);
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (targetStream as any).write = (...args: any[]): boolean => {
          if (streamAborted) return false;
          try {
            const ok: boolean = (origWrite as (...a: any[]) => boolean)(...args);
            if (!ok) needsDrain = true;
            return ok;
          } catch {
            streamAborted = true;
            return false;
          }
        };

        targetStream.on('drain', () => { needsDrain = false; });
        targetStream.on('close', () => { streamAborted = true; });
        targetStream.on('error', () => { streamAborted = true; });

        const rssThreshold = this._backpressureThreshold;

        drainFn = async () => {
          if (streamAborted) throw new Error('Client disconnected — output stream closed mid-export');

          // Primary: event-driven drain loop.
          // write() returning false is the authoritative Node.js backpressure signal.
          // Loop because after each TCP drain the archiver may immediately re-fill the stream.
          while (needsDrain && !streamAborted) {
            await new Promise<void>((resolve, reject) => {
              if (!needsDrain || streamAborted) { resolve(); return; }

              const cleanup = () => {
                targetStream.removeListener('drain', onDrain);
                targetStream.removeListener('close', onAbort);
                targetStream.removeListener('error', onAbort);
              };
              const onDrain = () => { cleanup(); resolve(); };
              const onAbort = () => { cleanup(); reject(new Error('Client disconnected — output stream closed mid-export')); };

              targetStream.on('drain', onDrain);
              targetStream.on('close', onAbort);
              targetStream.on('error', onAbort);

              const t = setTimeout(() => { cleanup(); resolve(); }, 30_000);
              if (typeof (t as NodeJS.Timeout).unref === 'function') (t as NodeJS.Timeout).unref();
            });

            if (!streamAborted) await new Promise<void>(r => setImmediate(r));
          }

          // Fallback: RSS-based polling.
          // When the output stream is behind a reverse proxy (nginx etc.), write() always
          // returns true (proxy accepts data instantly) so the drain loop above never fires.
          // Data still accumulates in Node.js's outputData buffer, growing process RSS.
          // Polling RSS directly works regardless of proxy topology.
          if (streamAborted) throw new Error('Client disconnected — output stream closed mid-export');
          const started = Date.now();
          while (process.memoryUsage().rss > rssThreshold && !streamAborted) {
            if (Date.now() - started > 30_000) break;
            await new Promise<void>(r => setTimeout(r, 200));
          }
          if (streamAborted) throw new Error('Client disconnected — output stream closed mid-export');
        };
      } else {
        // run() to file: no stream reference available, use RSS-only polling.
        // On slow storage (container overlay, network PVC) the archiver buffer
        // accumulates the same way as the pipe() case. RSS drops as disk writes
        // complete and Node.js releases the buffers.
        const rssThreshold = this._backpressureThreshold;
        drainFn = async () => {
          if (process.memoryUsage().rss <= rssThreshold) return;
          const started = Date.now();
          while (process.memoryUsage().rss > rssThreshold) {
            if (Date.now() - started > 30_000) break;
            await new Promise<void>(r => setTimeout(r, 200));
          }
        };
      }

      workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
        ...workbookTarget,
        useStyles       : true,
        useSharedStrings: false,
        // compress=false: skip ZIP deflate — dramatically faster streaming, slightly larger file
        zip             : this._compress ? undefined : { zlib: { level: 0 } },
      // ExcelJS streaming types omit the `zip` option — cast to silence the error
      } as unknown as ExcelJS.stream.xlsx.WorkbookWriterOptions);

      for (const sheetCfg of sheets) {
        const { sheetNames, skippedRows } = await this._executeSheet(connection, workbook, sheetCfg, progressCtx, drainFn);
        allSheets.push(...sheetNames);
        totalSkipped += skippedRows;
      }

      await workbook.commit();
      return { success: true, sheets: allSheets, skippedRows: totalSkipped };

    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error(`[OracleSqlToExcel:ERROR] _execute FAILED — ${msg}`);
      // End the stream on error to prevent hangs in toBuffer() / pipe() callers
      if (workbookTarget.stream && !workbookTarget.stream.writableEnded) {
        try { workbookTarget.stream.end(); } catch (_) {}
      }
      return { success: false, error: msg, sheets: allSheets, skippedRows: totalSkipped };
    } finally {
      if (connection) await connection.close().catch(() => {});
    }
  }

  // ── Terminal methods ───────────────────────────────────────────────────────

  /**
   * Execute and write the workbook to an `.xlsx` file on disk.
   * Deletes the partial file automatically if an error occurs mid-export.
   */
  async run(): Promise<RunResult | MultiRunResult> {
    if (this._files.length > 0) {
      if (this._files.length === 1) {
        return this._executeFileConfig(this._files[0]);
      }
      // Multiple .file() calls: run each in sequence, merge into one MultiRunResult
      const allFiles  : FileSegment[] = [];
      const allSheets : string[]      = [];
      let totalSkipped = 0;
      let success      = true;
      let firstError   : string | undefined;
      for (const fileCfg of this._files) {
        const r = await this._executeFileConfig(fileCfg);
        allFiles.push(...r.files);
        allSheets.push(...r.sheets);
        totalSkipped += r.skippedRows;
        if (!r.success) { success = false; firstError = firstError ?? r.error; }
      }
      return { success, files: allFiles, sheets: allSheets, skippedRows: totalSkipped, ...(firstError ? { error: firstError } : {}) };
    }
    const file   = path.join(this._outputDir, `${this._filePrefix}.xlsx`);
    const result = await this._execute({ filename: file });
    if (!result.success) {
      fs.promises.unlink(file).catch(() => {});
    }
    return { ...result, file };
  }

  /**
   * Execute and stream the workbook directly to a writable stream.
   * No temporary file is created — the Oracle query, workbook generation, and
   * data transfer are all streamed end-to-end.
   *
   * @example
   * res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
   * res.setHeader('Content-Disposition', 'attachment; filename="report.xlsx"');
   * await OracleSqlToExcel().sheet('Data', s => s.sql(SQL).columns(COLS)).pipe(res);
   */
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

  /**
   * Execute and return the workbook as an in-memory `Buffer`.
   * Useful for S3 uploads, email attachments, or storing as a DB BLOB
   * without writing to disk.
   *
   * @example
   * const { buffer } = await OracleSqlToExcel().sheet('Data', s => s.sql(SQL).columns(COLS)).toBuffer();
   * await s3.putObject({ Body: buffer, Key: 'report.xlsx' }).promise();
   */
  async toBuffer(): Promise<BufferResult> {
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
    const chunks: Buffer[] = [];
    const pass             = new PassThrough();
    pass.on('data', (chunk: Buffer) => chunks.push(chunk));

    const finishPromise = new Promise<void>((resolve, reject) => {
      pass.on('finish', resolve);
      pass.on('error', reject);
    });

    const result = await this._execute({ stream: pass });

    // _execute ends the stream on error; this is a defensive fallback
    if (!pass.writableEnded) pass.end();

    // Swallow stream errors — result.success already captures the failure state
    await finishPromise.catch(() => {});
    return { ...result, buffer: result.success ? Buffer.concat(chunks) : Buffer.alloc(0) };
  }
}

// ── Factory function ──────────────────────────────────────────────────────────

/**
 * Create a new {@link OracleSqlToExcelBuilder} instance.
 */
function OracleSqlToExcel(): OracleSqlToExcelBuilder {
  return new OracleSqlToExcelBuilder();
}

export { OracleSqlToExcel };
export default OracleSqlToExcel;

// ── CSV types ─────────────────────────────────────────────────────────────────

export interface CsvResult {
  success    : boolean;
  rowsWritten: number;
  error?     : string;
}

export interface CsvRunResult extends CsvResult {
  /** Absolute path to the written `.csv` file. */
  file: string;
}

export interface CsvBufferResult extends CsvResult {
  /** In-memory CSV buffer. Empty `Buffer` when `success` is `false`. */
  buffer: Buffer;
}

// ── CSV helpers ───────────────────────────────────────────────────────────────

/** @private */
function escapeCsvField(value: string, separator: string): string {
  if (value.includes(separator) || value.includes('"') || value.includes('\n') || value.includes('\r')) {
    return '"' + value.replace(/"/g, '""') + '"';
  }
  return value;
}

// ── OracleSqlToCsvBuilder ─────────────────────────────────────────────────────

/**
 * Fluent builder that streams Oracle SQL results directly into a `.csv` file or stream.
 *
 * CSV streaming writes each row directly to the output without intermediate buffering —
 * no ZIP archiver, no ExcelJS workbook. Memory usage is O(fetchSize × row_size) regardless
 * of total row count, making it ideal for very large exports (1M+ rows).
 *
 * Configure the query via `.sql()`, optional columns via `.columns()`, then call a terminal method.
 *
 * @example
 * // Stream to Express response
 * res.setHeader('Content-Type', 'text/csv; charset=utf-8');
 * res.setHeader('Content-Disposition', 'attachment; filename="report.csv"');
 * await OracleSqlToCsv()
 *   .connectionFactory(() => pool.getConnection())
 *   .sql('SELECT CODE, NAME, AMOUNT FROM BIG_TABLE')
 *   .columns([{ key: 'CODE', header: 'Code' }, { key: 'NAME', header: 'Name' }, { key: 'AMOUNT', header: 'Amount' }])
 *   .pipe(res);
 *
 * @example
 * // Write to file
 * const { file } = await OracleSqlToCsv()
 *   .connectionFactory(() => pool.getConnection())
 *   .sql('SELECT * FROM BIG_TABLE')
 *   .run('/tmp/export.csv');
 */
class OracleSqlToCsvBuilder {
  /** @private */ _connectionFactory : (() => Promise<OracleConnection>) | null;
  /** @private */ _sql               : string;
  /** @private */ _param             : Record<string, unknown>;
  /** @private */ _executeOptions    : Record<string, unknown>;
  /** @private */ _columns           : Pick<ColumnDef, 'key' | 'header'>[];
  /** @private */ _fetchSize         : number;
  /** @private */ _separator         : string;
  /** @private */ _withBom           : boolean;
  /** @private */ _onProgressCb      : ((info: { rowsWritten: number }) => void) | null;

  constructor() {
    this._connectionFactory = null;
    this._sql               = '';
    this._param             = {};
    this._executeOptions    = {};
    this._columns           = [];
    this._fetchSize         = 50_000;
    this._separator         = ',';
    this._withBom           = true;
    this._onProgressCb      = null;
  }

  /**
   * Factory function that returns a new Oracle connection. Called once per export.
   * @example .connectionFactory(() => pool.getConnection())
   */
  connectionFactory(fn: () => Promise<OracleConnection>): this { this._connectionFactory = fn; return this; }

  /**
   * Oracle SQL query with optional bind parameters and execute options.
   * @example .sql('SELECT * FROM T WHERE CODE = :code', { code: '019' })
   */
  sql(
    query          : string,
    param          : Record<string, unknown> = {},
    executeOptions : Record<string, unknown> = {}
  ): this {
    this._sql            = query;
    this._param          = param;
    this._executeOptions = executeOptions;
    return this;
  }

  /**
   * Column definitions. Only `key` and `header` are used for CSV (no type, style, or format).
   * Omit to auto-detect columns from Oracle metadata in their SELECT order.
   */
  columns(value: Pick<ColumnDef, 'key' | 'header'>[]): this { this._columns = value; return this; }

  /**
   * Rows fetched from Oracle per round-trip. Default: `50_000`.
   */
  fetchSize(value: number): this { this._fetchSize = value; return this; }

  /**
   * CSV field separator. Default: `','`.
   * @example .separator(';')  // European Excel
   * @example .separator('\t') // TSV
   */
  separator(value: string): this { this._separator = value; return this; }

  /**
   * Prepend a UTF-8 BOM (`﻿`) so Windows Excel opens the file with correct encoding.
   * Default: `true`.
   */
  withBom(value = true): this { this._withBom = value; return this; }

  /**
   * Callback invoked after each fetch batch is written.
   * @example .onProgress(({ rowsWritten }) => console.log(rowsWritten))
   */
  onProgress(cb: (info: { rowsWritten: number }) => void): this { this._onProgressCb = cb; return this; }

  // ── Internal ────────────────────────────────────────────────────────────────

  /** @private */
  async _execute(stream: Writable): Promise<CsvResult> {
    let connection : OracleConnection | null = null;
    let resultSet  : OracleResultSet  | null = null;
    let rowsWritten = 0;

    const write = (data: string): void => { stream.write(data); };

    try {
      if (!this._connectionFactory) throw new Error('No connection factory set.');
      if (!this._sql?.trim())        throw new Error('No SQL query set.');

      connection = await this._connectionFactory();

      const execResult = await connection.execute(this._sql, this._param, {
        autoCommit    : true,
        ...this._executeOptions,
        outFormat     : OUT_FORMAT_OBJECT,
        resultSet     : true,
        fetchArraySize: this._fetchSize,
      });

      resultSet = execResult.resultSet ?? null;

      // Resolve columns: explicit definition or Oracle metadata
      let cols = this._columns.length > 0
        ? this._columns
        : (execResult.metaData ?? []).map((m) => ({ key: m.name, header: m.name }));

      if (cols.length === 0) {
        // No metadata available — will resolve from first row
        let rows = await resultSet!.getRows(1);
        if (rows.length > 0) {
          cols = Object.keys(rows[0]).map((k) => ({ key: k, header: k }));
          // Write BOM + header before processing this pre-fetched row
          if (this._withBom) write('﻿');
          write(cols.map((c) => escapeCsvField(c.header ?? c.key, this._separator)).join(this._separator) + '\n');
          for (const row of rows) {
            write(cols.map((c) => escapeCsvField(String(row[c.key] ?? ''), this._separator)).join(this._separator) + '\n');
            rowsWritten++;
          }
          if (this._onProgressCb) this._onProgressCb({ rowsWritten });
          rows = await resultSet!.getRows(this._fetchSize);
          while (rows.length > 0) {
            const lines = rows.map((row) =>
              cols.map((c) => escapeCsvField(String(row[c.key] ?? ''), this._separator)).join(this._separator) + '\n'
            ).join('');
            write(lines);
            rowsWritten += rows.length;
            if (this._onProgressCb) this._onProgressCb({ rowsWritten });
            rows = await resultSet!.getRows(this._fetchSize);
          }
          return { success: true, rowsWritten };
        }
        return { success: true, rowsWritten: 0 };
      }

      // Write BOM + header
      if (this._withBom) write('﻿');
      write(cols.map((c) => escapeCsvField(c.header ?? c.key, this._separator)).join(this._separator) + '\n');

      // Stream rows batch by batch — no intermediate accumulation
      let rows = await resultSet!.getRows(this._fetchSize);
      while (rows.length > 0) {
        const lines = rows.map((row) =>
          cols.map((c) => escapeCsvField(String(row[c.key] ?? ''), this._separator)).join(this._separator) + '\n'
        ).join('');
        write(lines);
        rowsWritten += rows.length;
        if (this._onProgressCb) this._onProgressCb({ rowsWritten });
        rows = await resultSet!.getRows(this._fetchSize);
      }

      return { success: true, rowsWritten };
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error(`[OracleSqlToCsv:ERROR] ${msg}`);
      if (!stream.writableEnded) stream.end();
      return { success: false, rowsWritten, error: msg };
    } finally {
      if (resultSet)  await resultSet.close().catch(() => {});
      if (connection) await connection.close().catch(() => {});
    }
  }

  // ── Terminal methods ───────────────────────────────────────────────────────

  /**
   * Write CSV to a file at `filepath`.
   * @param filepath - Absolute or relative path including filename and extension (e.g. `'/tmp/report.csv'`).
   */
  async run(filepath: string): Promise<CsvRunResult> {
    const stream = fs.createWriteStream(filepath);
    const done   = new Promise<void>((resolve, reject) => {
      stream.on('finish', resolve);
      stream.on('error', reject);
    });
    const result = await this._execute(stream);
    if (!stream.writableEnded) stream.end();
    await done.catch(() => {});
    if (!result.success) fs.promises.unlink(filepath).catch(() => {});
    return { ...result, file: filepath };
  }

  /**
   * Stream CSV directly to any Writable (e.g. Express `res`).
   *
   * @example
   * res.setHeader('Content-Type', 'text/csv; charset=utf-8');
   * res.setHeader('Content-Disposition', 'attachment; filename="report.csv"');
   * await OracleSqlToCsv().connectionFactory(...).sql(SQL).pipe(res);
   */
  async pipe(writableStream: Writable): Promise<CsvResult> {
    return this._execute(writableStream);
  }

  /**
   * Return the entire CSV as an in-memory `Buffer`.
   * For large datasets prefer `.run()` or `.pipe()` to avoid holding all data in memory.
   */
  async toBuffer(): Promise<CsvBufferResult> {
    const chunks: Buffer[] = [];
    const pass             = new PassThrough();
    pass.on('data', (chunk: Buffer) => chunks.push(chunk));
    const done = new Promise<void>((resolve, reject) => {
      pass.on('finish', resolve);
      pass.on('error', reject);
    });
    const result = await this._execute(pass);
    if (!pass.writableEnded) pass.end();
    await done.catch(() => {});
    return { ...result, buffer: result.success ? Buffer.concat(chunks) : Buffer.alloc(0) };
  }
}

// ── CSV Factory ───────────────────────────────────────────────────────────────

/**
 * Create a new {@link OracleSqlToCsvBuilder} instance.
 */
function OracleSqlToCsv(): OracleSqlToCsvBuilder {
  return new OracleSqlToCsvBuilder();
}

export { OracleSqlToCsv };
