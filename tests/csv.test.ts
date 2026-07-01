import fs from 'fs';
import os from 'os';
import path from 'path';
import { OracleSqlToCsv } from '../src/index';
import { createStreamConn } from './helpers';

// ── helpers ───────────────────────────────────────────────────────────────────

const COLS = [
  { key: 'ID',   header: 'ID'   },
  { key: 'NAME', header: 'Name' },
];

const ROWS = [
  { ID: 1, NAME: 'Alice' },
  { ID: 2, NAME: 'Bob'   },
];

function makeConn(rows: Record<string, unknown>[]) {
  return () => Promise.resolve(createStreamConn(rows, COLS.map((c) => ({ name: c.key }))));
}

function parseCsv(csv: string, separator = ','): string[][] {
  return csv
    .split('\n')
    .filter((line) => line.trim() !== '')
    .map((line) => line.split(separator).map((f) => f.replace(/^"|"$/g, '').replace(/""/g, '"')));
}

// ── toBuffer ──────────────────────────────────────────────────────────────────

describe('OracleSqlToCsv — toBuffer()', () => {
  it('returns success=true and non-empty buffer', async () => {
    const { success, buffer, rowsWritten } = await OracleSqlToCsv()
      .connectionFactory(makeConn(ROWS))
      .sql('SELECT * FROM T')
      .columns(COLS)
      .toBuffer();

    expect(success).toBe(true);
    expect(buffer.length).toBeGreaterThan(0);
    expect(rowsWritten).toBe(2);
  });

  it('writes BOM + correct headers and data rows', async () => {
    const { buffer } = await OracleSqlToCsv()
      .connectionFactory(makeConn(ROWS))
      .sql('SELECT * FROM T')
      .columns(COLS)
      .toBuffer();

    const text = buffer.toString('utf8');
    expect(text.charCodeAt(0)).toBe(0xFEFF); // BOM
    const rows = parseCsv(text.replace(/^﻿/, ''));
    expect(rows[0]).toEqual(['ID', 'Name']);
    expect(rows[1]).toEqual(['1', 'Alice']);
    expect(rows[2]).toEqual(['2', 'Bob']);
  });

  it('withBom(false) omits BOM', async () => {
    const { buffer } = await OracleSqlToCsv()
      .connectionFactory(makeConn(ROWS))
      .sql('SELECT * FROM T')
      .columns(COLS)
      .withBom(false)
      .toBuffer();

    const text = buffer.toString('utf8');
    expect(text.charCodeAt(0)).not.toBe(0xFEFF);
    expect(text.startsWith('ID')).toBe(true);
  });

  it('respects custom separator', async () => {
    const { buffer } = await OracleSqlToCsv()
      .connectionFactory(makeConn(ROWS))
      .sql('SELECT * FROM T')
      .columns(COLS)
      .separator(';')
      .withBom(false)
      .toBuffer();

    const text = buffer.toString('utf8');
    const rows = parseCsv(text, ';');
    expect(rows[0]).toEqual(['ID', 'Name']);
    expect(rows[1]).toEqual(['1', 'Alice']);
  });

  it('auto-detects columns from metadata when .columns() omitted', async () => {
    const { buffer, rowsWritten } = await OracleSqlToCsv()
      .connectionFactory(makeConn(ROWS))
      .sql('SELECT * FROM T')
      .toBuffer();

    expect(rowsWritten).toBe(2);
    const text = buffer.toString('utf8').replace(/^﻿/, '');
    const rows = parseCsv(text);
    expect(rows[0]).toEqual(['ID', 'NAME']);
  });

  it('returns success=false on connection error', async () => {
    const { success, buffer, rowsWritten, error } = await OracleSqlToCsv()
      .connectionFactory(() => Promise.reject(new Error('DB down')))
      .sql('SELECT 1 FROM DUAL')
      .toBuffer();

    expect(success).toBe(false);
    expect(buffer.length).toBe(0);
    expect(rowsWritten).toBe(0);
    expect(error).toMatch(/DB down/);
  });

  it('calls onProgress after each batch', async () => {
    const calls: number[] = [];
    await OracleSqlToCsv()
      .connectionFactory(makeConn(ROWS))
      .sql('SELECT * FROM T')
      .columns(COLS)
      .fetchSize(1)
      .onProgress(({ rowsWritten }) => calls.push(rowsWritten))
      .toBuffer();

    expect(calls.length).toBeGreaterThan(0);
    expect(calls[calls.length - 1]).toBe(2);
  });

  it('escapes fields containing the separator', async () => {
    const rows = [{ ID: 1, NAME: 'Smith, Jr.' }];
    const { buffer } = await OracleSqlToCsv()
      .connectionFactory(makeConn(rows))
      .sql('SELECT * FROM T')
      .columns(COLS)
      .withBom(false)
      .toBuffer();

    const text = buffer.toString('utf8');
    expect(text).toContain('"Smith, Jr."');
  });
});

// ── run() ─────────────────────────────────────────────────────────────────────

describe('OracleSqlToCsv — run()', () => {
  let tmpDir: string;

  beforeEach(() => { tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'csv-test-')); });
  afterEach(()  => { fs.rmSync(tmpDir, { recursive: true, force: true }); });

  it('writes file to disk and returns path + rowsWritten', async () => {
    const filepath = path.join(tmpDir, 'out.csv');
    const _r = await OracleSqlToCsv()
      .connectionFactory(makeConn(ROWS))
      .sql('SELECT * FROM T')
      .columns(COLS)
      .run(filepath);
    const { success, rowsWritten } = _r;
    const file = 'file' in _r ? _r.file : undefined;

    expect(success).toBe(true);
    expect(file).toBe(filepath);
    expect(rowsWritten).toBe(2);
    expect(fs.existsSync(filepath)).toBe(true);
  });

  it('deletes partial file on error', async () => {
    const filepath = path.join(tmpDir, 'bad.csv');
    const { success } = await OracleSqlToCsv()
      .connectionFactory(() => Promise.reject(new Error('fail')))
      .sql('SELECT 1 FROM DUAL')
      .run(filepath);

    expect(success).toBe(false);
    expect(fs.existsSync(filepath)).toBe(false);
  });
});

// ── pipe() ────────────────────────────────────────────────────────────────────

describe('OracleSqlToCsv — pipe()', () => {
  it('streams CSV bytes to writable', async () => {
    const { PassThrough } = await import('stream');
    const chunks: Buffer[] = [];
    const pass = new PassThrough();
    pass.on('data', (c: Buffer) => chunks.push(c));
    const done = new Promise<void>((res) => pass.on('finish', res));

    const result = await OracleSqlToCsv()
      .connectionFactory(makeConn(ROWS))
      .sql('SELECT * FROM T')
      .columns(COLS)
      .withBom(false)
      .pipe(pass);

    pass.end();
    await done;

    expect(result.success).toBe(true);
    const text = Buffer.concat(chunks).toString('utf8');
    const rows = parseCsv(text);
    expect(rows[0]).toEqual(['ID', 'Name']);
    expect(rows).toHaveLength(3); // header + 2 data rows
  });
});

// ── transform ─────────────────────────────────────────────────────────────────

describe('OracleSqlToCsv — transform', () => {
  async function csvToString(
    builder: ReturnType<typeof OracleSqlToCsv>
  ): Promise<string> {
    const { buffer } = await builder.withBom(false).toBuffer();
    return buffer.toString('utf8');
  }

  it('per-column transform applied to raw Oracle value', async () => {
    const rows = [{ STATUS: 'Y' }, { STATUS: 'N' }];
    const cols = [{
      key      : 'STATUS',
      transform: (val: unknown) => val === 'Y' ? 'Active' : 'Inactive',
    }];

    const csv = await csvToString(
      OracleSqlToCsv()
        .connectionFactory(() => Promise.resolve(createStreamConn(rows, [{ name: 'STATUS' }])))
        .sql('SELECT STATUS FROM T')
        .columns(cols)
    );

    const dataRows = csv.trim().split('\n').slice(1);
    expect(dataRows[0]).toBe('Active');
    expect(dataRows[1]).toBe('Inactive');
  });

  it('per-row transform appends computed column', async () => {
    const rows = [{ A: 'foo', B: 'bar' }];
    const cols = [{ key: 'A' }, { key: 'B' }, { key: 'C' }];

    const csv = await csvToString(
      OracleSqlToCsv()
        .connectionFactory(() => Promise.resolve(createStreamConn(rows, [{ name: 'A' }, { name: 'B' }])))
        .sql('SELECT A, B FROM T')
        .columns(cols)
        .transform((row) => ({ ...row, C: `${row['A']}-${row['B']}` }))
    );

    expect(csv).toContain('foo-bar');
  });

  it('per-row transform receives row after per-column transforms', async () => {
    const rows = [{ FLAG: 'Y' }];
    const cols = [
      { key: 'FLAG', transform: (val: unknown) => val === 'Y' ? 'yes' : 'no' },
      { key: 'LABEL' },
    ];

    const csv = await csvToString(
      OracleSqlToCsv()
        .connectionFactory(() => Promise.resolve(createStreamConn(rows, [{ name: 'FLAG' }])))
        .sql('SELECT FLAG FROM T')
        .columns(cols)
        .transform((row) => ({ ...row, LABEL: `flag=${row['FLAG']}` }))
    );

    const dataRows = csv.trim().split('\n').slice(1);
    expect(dataRows[0]).toBe('yes,flag=yes');
  });

  it('warns once in dev when CSV transform returns Promise', async () => {
    const origEnv = process.env.NODE_ENV;
    const warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
    try {
      process.env.NODE_ENV = 'development';
      const rows = [{ A: '1' }, { A: '2' }, { A: '3' }];
      const cols = [{ key: 'A', transform: async (val: unknown) => String(val).toUpperCase() }];

      await OracleSqlToCsv()
        .connectionFactory(() => Promise.resolve(createStreamConn(rows, [{ name: 'A' }])))
        .sql('SELECT A FROM T')
        .columns(cols)
        .withBom(false)
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

  it('async CSV transform produces correct values', async () => {
    const rows = [{ V: 'x' }, { V: 'y' }];
    const cols = [{ key: 'V', transform: async (val: unknown) => String(val).toUpperCase() }];

    const csv = await csvToString(
      OracleSqlToCsv()
        .connectionFactory(() => Promise.resolve(createStreamConn(rows, [{ name: 'V' }])))
        .sql('SELECT V FROM T')
        .columns(cols)
    );

    const dataRows = csv.trim().split('\n').slice(1);
    expect(dataRows[0]).toBe('X');
    expect(dataRows[1]).toBe('Y');
  });
});
