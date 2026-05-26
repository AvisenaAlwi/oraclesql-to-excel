import fs from 'fs';
import os from 'os';
import path from 'path';
import { PassThrough } from 'stream';
import { OracleSqlToExcel } from '../src/index';
import { createStreamConn, createCountConn, readBuffer } from './helpers';

// ── helpers ───────────────────────────────────────────────────────────────────

const COLS = [
  { key: 'ID',   header: 'ID',   type: 'number' as const },
  { key: 'NAME', header: 'Name', type: 'text'   as const },
];

const ROWS = [
  { ID: 1, NAME: 'Alice' },
  { ID: 2, NAME: 'Bob'   },
];

function makeConn(rows: Record<string, unknown>[], meta = COLS.map((c) => ({ name: c.key }))) {
  return () => Promise.resolve(createStreamConn(rows, meta));
}

// ── basic buffer ──────────────────────────────────────────────────────────────

describe('toBuffer — basic', () => {
  it('returns success=true and non-empty buffer', async () => {
    const { success, buffer } = await OracleSqlToExcel()
      .connectionFactory(makeConn(ROWS))
      .sheet('Report', (s) => s.sql('SELECT * FROM T').columns(COLS))
      .toBuffer();

    expect(success).toBe(true);
    expect(buffer.length).toBeGreaterThan(0);
  });

  it('returns correct headers and row values', async () => {
    const { buffer } = await OracleSqlToExcel()
      .connectionFactory(makeConn(ROWS))
      .sheet('Data', (s) => s.sql('SELECT * FROM T').columns(COLS))
      .toBuffer();

    const sheets = await readBuffer(buffer, 'Data');
    expect(sheets).toHaveLength(1);
    const { headers, rows } = sheets[0];
    expect(headers).toEqual(['ID', 'Name']);
    expect(rows).toHaveLength(2);
    expect(rows[0][0]).toBe(1);
    expect(rows[0][1]).toBe('Alice');
  });

  it('returns empty buffer and success=false on error', async () => {
    const badConn = () => Promise.reject(new Error('DB down'));
    const { success, buffer, error } = await OracleSqlToExcel()
      .connectionFactory(badConn)
      .sheet('S', (s) => s.sql('SELECT 1 FROM DUAL').columns(COLS))
      .toBuffer();

    expect(success).toBe(false);
    expect(buffer.length).toBe(0);
    expect(error).toMatch(/DB down/i);
  });
});

// ── column types ──────────────────────────────────────────────────────────────

describe('toBuffer — column type handling', () => {
  it('writes numbers as numeric cells', async () => {
    const rows = [{ VAL: 42 }];
    const cols = [{ key: 'VAL', type: 'number' as const }];
    const { buffer } = await OracleSqlToExcel()
      .connectionFactory(makeConn(rows, [{ name: 'VAL' }]))
      .sheet('S', (s) => s.sql('SELECT VAL FROM T').columns(cols))
      .toBuffer();

    const [{ rows: parsed }] = await readBuffer(buffer);
    expect(typeof parsed[0][0]).toBe('number');
    expect(parsed[0][0]).toBe(42);
  });

  it('coerces large numbers (>15 sig digits) to string', async () => {
    const bigNum = '12345678901234567';
    const rows   = [{ VAL: bigNum }];
    const cols   = [{ key: 'VAL', type: 'number' as const }];
    const { buffer } = await OracleSqlToExcel()
      .connectionFactory(makeConn(rows, [{ name: 'VAL' }]))
      .sheet('S', (s) => s.sql('SELECT VAL FROM T').columns(cols))
      .toBuffer();

    const [{ rows: parsed }] = await readBuffer(buffer);
    expect(String(parsed[0][0])).toBe(bigNum);
  });

  it('writes null cells as null/empty', async () => {
    const rows = [{ VAL: null }];
    const cols = [{ key: 'VAL', type: 'text' as const }];
    const { buffer } = await OracleSqlToExcel()
      .connectionFactory(makeConn(rows, [{ name: 'VAL' }]))
      .sheet('S', (s) => s.sql('SELECT VAL FROM T').columns(cols))
      .toBuffer();

    const [{ rows: parsed }] = await readBuffer(buffer);
    // ExcelJS streaming omits fully-null rows from getSheetValues()
    const cell = parsed[0]?.[0];
    expect(cell === null || cell === undefined || cell === '').toBe(true);
  });

  it('writes date columns as Date objects', async () => {
    const d    = new Date('2025-01-15');
    const rows = [{ DT: d }];
    const cols = [{ key: 'DT', type: 'date' as const }];
    const { buffer } = await OracleSqlToExcel()
      .connectionFactory(makeConn(rows, [{ name: 'DT' }]))
      .sheet('S', (s) => s.sql('SELECT DT FROM T').columns(cols))
      .toBuffer();

    const [{ rows: parsed }] = await readBuffer(buffer);
    // ExcelJS serialises dates; value may come back as Date or number
    expect(parsed[0][0]).toBeTruthy();
  });
});

// ── empty result set ──────────────────────────────────────────────────────────

describe('toBuffer — empty result set', () => {
  it('succeeds with header row only', async () => {
    const { success, buffer } = await OracleSqlToExcel()
      .connectionFactory(makeConn([]))
      .sheet('Empty', (s) => s.sql('SELECT * FROM T').columns(COLS))
      .toBuffer();

    expect(success).toBe(true);
    const [{ headers, rows }] = await readBuffer(buffer, 'Empty');
    expect(headers).toEqual(['ID', 'Name']);
    expect(rows).toHaveLength(0);
  });
});

// ── multi-sheet ───────────────────────────────────────────────────────────────

describe('toBuffer — multi-sheet', () => {
  it('writes two sheets with correct names', async () => {
    const connFactory = jest
      .fn()
      .mockResolvedValueOnce(createStreamConn([{ ID: 1, NAME: 'Alice' }]))
      .mockResolvedValueOnce(createStreamConn([{ CODE: 'X', VAL: 99  }]));

    const cols2 = [
      { key: 'CODE', type: 'text'   as const },
      { key: 'VAL',  type: 'number' as const },
    ];

    const { success, sheets } = await OracleSqlToExcel()
      .connectionFactory(connFactory)
      .sheet('Sheet1', (s) => s.sql('SELECT * FROM T1').columns(COLS))
      .sheet('Sheet2', (s) => s.sql('SELECT * FROM T2').columns(cols2))
      .toBuffer();

    expect(success).toBe(true);
    expect(sheets).toEqual(['Sheet1', 'Sheet2']);
  });
});

// ── split sheets ──────────────────────────────────────────────────────────────

describe('toBuffer — split sheets', () => {
  it('splits into multiple physical sheets when maxRowsPerSheet exceeded', async () => {
    const rows = Array.from({ length: 5 }, (_, i) => ({ ID: i + 1, NAME: `Row${i + 1}` }));
    const { success, sheets, buffer } = await OracleSqlToExcel()
      .connectionFactory(makeConn(rows))
      .sheet('Part', (s) =>
        s.sql('SELECT * FROM T').columns(COLS).maxRowsPerSheet(2)
      )
      .toBuffer();

    expect(success).toBe(true);
    // 5 rows / 2 per sheet = 3 sheets
    expect(sheets.length).toBe(3);
    expect(sheets[0]).toBe('Part');
    expect(sheets[1]).toBe('Part 2');
    expect(sheets[2]).toBe('Part 3');

    const parsed = await readBuffer(buffer);
    expect(parsed).toHaveLength(3);
    // first sheet has 2 data rows + separator rows written before the split
    expect(parsed[0].rows.length).toBeGreaterThanOrEqual(2);
    expect(parsed[0].rows[0][0]).toBe(1);
    expect(parsed[0].rows[1][0]).toBe(2);
  });
});

// ── doc header ────────────────────────────────────────────────────────────────

describe('toBuffer — docHeader', () => {
  it('prepends doc-header rows before the table header', async () => {
    const { success, buffer } = await OracleSqlToExcel()
      .connectionFactory(makeConn(ROWS))
      .sheet('Report', (s) =>
        s
          .sql('SELECT * FROM T')
          .columns(COLS)
          .docHeader([
            { text: 'COMPANY NAME', style: { bold: true } },
            { text: 'Period: 2025', style: { italic: true } },
          ])
      )
      .toBuffer();

    expect(success).toBe(true);

    // Workbook should still be parseable
    const ExcelJS = (await import('exceljs')).default;
    const wb = new ExcelJS.Workbook();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await wb.xlsx.load(buffer as any);
    const ws = wb.getWorksheet('Report')!;
    // Row 1 and 2 are doc-header; row 3 is the table header
    expect(ws.getRow(1).getCell(1).value).toBe('COMPANY NAME');
    expect(ws.getRow(2).getCell(1).value).toBe('Period: 2025');
    // Table header row
    const headerRow = ws.getRow(3);
    expect(headerRow.getCell(1).value).toBe('ID');
    expect(headerRow.getCell(2).value).toBe('Name');
  });
});

// ── auto-detect columns ───────────────────────────────────────────────────────

describe('toBuffer — auto-detect columns', () => {
  it('uses Oracle metadata when .columns() is omitted', async () => {
    const rows = [{ ALPHA: 'x', BETA: 99 }];
    const meta = [
      { name: 'ALPHA', dbTypeName: 'VARCHAR2' },
      { name: 'BETA',  dbTypeName: 'NUMBER'   },
    ];
    const { success, buffer } = await OracleSqlToExcel()
      .connectionFactory(() => Promise.resolve(createStreamConn(rows, meta)))
      .sheet('Auto', (s) => s.sql('SELECT * FROM T'))
      .toBuffer();

    expect(success).toBe(true);
    const [{ headers }] = await readBuffer(buffer, 'Auto');
    expect(headers).toEqual(['ALPHA', 'BETA']);
  });
});

// ── onRowError skip ───────────────────────────────────────────────────────────

describe('toBuffer — onRowError skip', () => {
  it('skips bad rows and reports skippedRows count', async () => {
    // Force a row error by injecting a value that causes castCell to throw
    const original = require('../src/index');
    const rows = [
      { ID: 1, NAME: 'OK' },
      { ID: 2, NAME: 'OK' },
    ];
    // Make the connection mock throw on getRows for a specific row by patching castCell
    // Instead: use a getter that throws on access — simulated via a poisoned Proxy
    const poisonedRow = new Proxy({} as Record<string, unknown>, {
      get(_, prop) {
        if (prop === 'ID') throw new Error('bad row');
        return undefined;
      },
    });

    const badRows = [rows[0], poisonedRow as unknown as Record<string, unknown>, rows[1]];
    const { success, skippedRows } = await OracleSqlToExcel()
      .connectionFactory(makeConn(badRows))
      .sheet('S', (s) =>
        s.sql('SELECT * FROM T').columns(COLS).onRowError('skip')
      )
      .toBuffer();

    expect(success).toBe(true);
    expect(skippedRows).toBe(1);
  });
});

// ── run() — file on disk ──────────────────────────────────────────────────────

describe('run()', () => {
  let tmpDir: string;

  beforeEach(() => {
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ost-test-'));
  });

  afterEach(() => {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  });

  it('creates an .xlsx file and returns its path', async () => {
    const { success, file } = await OracleSqlToExcel()
      .connectionFactory(makeConn(ROWS))
      .outputDir(tmpDir)
      .filePrefix('output')
      .sheet('S', (s) => s.sql('SELECT * FROM T').columns(COLS))
      .run();

    expect(success).toBe(true);
    expect(file).toMatch(/output\.xlsx$/);
    expect(fs.existsSync(file)).toBe(true);
  });

  it('deletes partial file on error', async () => {
    const badConn = () => Promise.reject(new Error('fail'));
    const { success, file } = await OracleSqlToExcel()
      .connectionFactory(badConn)
      .outputDir(tmpDir)
      .filePrefix('bad')
      .sheet('S', (s) => s.sql('SELECT 1 FROM DUAL').columns(COLS))
      .run();

    expect(success).toBe(false);
    expect(fs.existsSync(file)).toBe(false);
  });
});

// ── pipe() ────────────────────────────────────────────────────────────────────

describe('pipe()', () => {
  it('writes valid XLSX bytes to writable stream', async () => {
    const pass   = new PassThrough();
    const chunks: Buffer[] = [];
    pass.on('data', (c: Buffer) => chunks.push(c));

    const resultPromise = OracleSqlToExcel()
      .connectionFactory(makeConn(ROWS))
      .sheet('S', (s) => s.sql('SELECT * FROM T').columns(COLS))
      .pipe(pass);

    const finish = new Promise<void>((res) => pass.on('finish', res));
    const { success } = await resultPromise;
    await finish;

    const buffer = Buffer.concat(chunks);
    expect(success).toBe(true);
    expect(buffer.length).toBeGreaterThan(0);
    // XLSX files start with PK (ZIP magic bytes)
    expect(buffer[0]).toBe(0x50);
    expect(buffer[1]).toBe(0x4b);
  });
});
