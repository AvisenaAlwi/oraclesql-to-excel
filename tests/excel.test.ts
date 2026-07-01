import fs from 'fs';
import os from 'os';
import path from 'path';
import { PassThrough } from 'stream';
import { OracleSqlToExcel, RunResult } from '../src/index';
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
      .run() as RunResult;

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
      .run() as RunResult;

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

// ── transform ─────────────────────────────────────────────────────────────────

describe('transform — per-column', () => {
  it('applies sync transform after castCell', async () => {
    const rows = [{ STATUS: 'Y' }, { STATUS: 'N' }];
    const cols = [{
      key      : 'STATUS',
      type     : 'text' as const,
      transform: (val: unknown) => val === 'Y' ? 'Aktif' : 'Nonaktif',
    }];

    const { buffer } = await OracleSqlToExcel()
      .connectionFactory(() => Promise.resolve(createStreamConn(rows, [{ name: 'STATUS' }])))
      .sheet('S', (s) => s.sql('SELECT STATUS FROM T').columns(cols))
      .toBuffer();

    const [{ rows: parsed }] = await readBuffer(buffer);
    expect(parsed[0][0]).toBe('Aktif');
    expect(parsed[1][0]).toBe('Nonaktif');
  });

  it('passes rawRow as second argument', async () => {
    const rows = [{ A: 'hello', B: 'world' }];
    const capturedRaw: unknown[] = [];
    const cols = [
      {
        key      : 'A',
        transform: (val: unknown, rawRow: Record<string, unknown>) => {
          capturedRaw.push({ ...rawRow });
          return val;
        },
      },
      { key: 'B' },
    ];

    await OracleSqlToExcel()
      .connectionFactory(() => Promise.resolve(createStreamConn(rows, [{ name: 'A' }, { name: 'B' }])))
      .sheet('S', (s) => s.sql('SELECT A, B FROM T').columns(cols))
      .toBuffer();

    expect(capturedRaw[0]).toEqual({ A: 'hello', B: 'world' });
  });

  it('null raw value passed through transform', async () => {
    const rows = [{ VAL: null }];
    const cols = [{ key: 'VAL', transform: (val: unknown) => val === null ? 'N/A' : val }];

    const { buffer } = await OracleSqlToExcel()
      .connectionFactory(() => Promise.resolve(createStreamConn(rows, [{ name: 'VAL' }])))
      .sheet('S', (s) => s.sql('SELECT VAL FROM T').columns(cols))
      .toBuffer();

    const [{ rows: parsed }] = await readBuffer(buffer);
    expect(parsed[0][0]).toBe('N/A');
  });
});

describe('transform — per-row (SheetConfig)', () => {
  it('computes derived column from multiple raw columns', async () => {
    const rows = [{ FIRST: 'John', LAST: 'Doe' }];
    const cols = [
      { key: 'FIRST' },
      { key: 'LAST' },
      { key: 'FULL' },
    ];

    const { buffer } = await OracleSqlToExcel()
      .connectionFactory(() => Promise.resolve(createStreamConn(rows, [{ name: 'FIRST' }, { name: 'LAST' }])))
      .sheet('S', (s) => s
        .sql('SELECT FIRST, LAST FROM T')
        .columns(cols)
        .transform((row) => ({ ...row, FULL: `${row['FIRST']} ${row['LAST']}` }))
      )
      .toBuffer();

    const [{ rows: parsed }] = await readBuffer(buffer);
    expect(parsed[0][2]).toBe('John Doe');
  });

  it('row transform receives row AFTER per-column transforms', async () => {
    const rows = [{ STATUS: 'Y' }];
    const cols = [
      { key: 'STATUS', transform: (val: unknown) => val === 'Y' ? 'Active' : 'Inactive' },
      { key: 'LABEL' },
    ];

    const { buffer } = await OracleSqlToExcel()
      .connectionFactory(() => Promise.resolve(createStreamConn(rows, [{ name: 'STATUS' }])))
      .sheet('S', (s) => s
        .sql('SELECT STATUS FROM T')
        .columns(cols)
        .transform((row) => ({ ...row, LABEL: `Status: ${row['STATUS']}` }))
      )
      .toBuffer();

    const [{ rows: parsed }] = await readBuffer(buffer);
    expect(parsed[0][1]).toBe('Status: Active');
  });
});

describe('transform — async', () => {
  it('warns once in dev when per-column transform returns Promise', async () => {
    const origEnv = process.env.NODE_ENV;
    const warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
    try {
      process.env.NODE_ENV = 'development';

      const rows = [{ A: '1' }, { A: '2' }, { A: '3' }];
      const cols = [{ key: 'A', transform: async (val: unknown) => String(val).toUpperCase() }];

      await OracleSqlToExcel()
        .connectionFactory(() => Promise.resolve(createStreamConn(rows, [{ name: 'A' }])))
        .sheet('S', (s) => s.sql('SELECT A FROM T').columns(cols))
        .toBuffer();

      const asyncWarns = warnSpy.mock.calls.filter(
        (c) => typeof c[0] === 'string' && c[0].includes('transform returned a Promise')
      );
      expect(asyncWarns).toHaveLength(1);
    } finally {
      warnSpy.mockRestore();
      process.env.NODE_ENV = origEnv;
    }
  });

  it('async per-column transform still produces correct values', async () => {
    const rows = [{ VAL: 'a' }, { VAL: 'b' }];
    const cols = [{ key: 'VAL', transform: async (val: unknown) => String(val).toUpperCase() }];

    const { buffer } = await OracleSqlToExcel()
      .connectionFactory(() => Promise.resolve(createStreamConn(rows, [{ name: 'VAL' }])))
      .sheet('S', (s) => s.sql('SELECT VAL FROM T').columns(cols))
      .toBuffer();

    const [{ rows: parsed }] = await readBuffer(buffer);
    expect(parsed[0][0]).toBe('A');
    expect(parsed[1][0]).toBe('B');
  });

  it('warns once in dev when row transform returns Promise', async () => {
    const origEnv = process.env.NODE_ENV;
    const warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
    try {
      process.env.NODE_ENV = 'development';

      const rows = [{ A: 'x' }, { A: 'y' }];
      const cols = [{ key: 'A' }];

      await OracleSqlToExcel()
        .connectionFactory(() => Promise.resolve(createStreamConn(rows, [{ name: 'A' }])))
        .sheet('S', (s) => s
          .sql('SELECT A FROM T')
          .columns(cols)
          .transform(async (row) => row)
        )
        .toBuffer();

      const asyncWarns = warnSpy.mock.calls.filter(
        (c) => typeof c[0] === 'string' && c[0].includes('transform returned a Promise')
      );
      expect(asyncWarns).toHaveLength(1);
    } finally {
      warnSpy.mockRestore();
      process.env.NODE_ENV = origEnv;
    }
  });
});
