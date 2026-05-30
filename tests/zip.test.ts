import fs from 'fs';
import os from 'os';
import path from 'path';
import { PassThrough } from 'stream';
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
    // rawRows is 0-indexed (plain push array); each element's inner values are 1-indexed (row.values).
    // rawRows[0] is the first Excel row = showTotalRows line; [1] is the first cell value.
    const showTotalRowsText = String(rawRows[0]?.[1] ?? '');
    // Must show "4" as start, not "1"
    expect(showTotalRowsText).toMatch(/4/);
    expect(showTotalRowsText).toMatch(/5/);
    expect(showTotalRowsText).toMatch(/5 total/i);
    expect(showTotalRowsText).not.toMatch(/^Showing rows 1/);

    fs.rmSync(tmpDir, { recursive: true });
  });
});

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
    expect(sheet1.rows).toHaveLength(5); // 3 data rows + empty row + next-file note
    expect(sheet1.rows[0][0]).toBe(1);
    expect(sheet1.rows[2][0]).toBe(3);

    const [sheet2] = await readBuffer(entries[1].data, 'Data');
    expect(sheet2.rows).toHaveLength(3); // prev-file note + 2 data rows
    expect(sheet2.rows[1][0]).toBe(4); // data starts at index 1 (after note)
    expect(sheet2.rows[2][0]).toBe(5);
  });

  it('cross-file notes contain correct filenames', async () => {
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
    const [sheet2] = await readBuffer(entries[1].data, 'Data');

    // last row of file 1 = "Next data available on file: rep_2.xlsx"
    const nextNote = String(sheet1.rows[sheet1.rows.length - 1][0] ?? '');
    expect(nextNote).toContain('rep_2.xlsx');

    // first row of file 2 = "Previous data on file: rep_1.xlsx"
    const prevNote = String(sheet2.rows[0][0] ?? '');
    expect(prevNote).toContain('rep_1.xlsx');
  });

  it('maxRowsPerSheet larger than maxRowsPerFile does not create empty split sheets', async () => {
    // maxRowsPerSheet=1_000_000 (default) >> maxRowsPerFile=3 — sheet should NOT split
    const { buffer } = await OracleSqlToExcel()
      .connectionFactory(factory)
      .file('rep', (f) => f
        .maxRowsPerFile(3)
        .sheet('Data', (s) => s.sql('SELECT * FROM T').columns(COLS))
      )
      .asZip()
      .toBuffer();

    const entries = readZipBuffer(buffer);
    expect(entries).toHaveLength(2);

    // each entry should have exactly 1 worksheet named 'Data'
    const [file1Sheet] = await readBuffer(entries[0].data, 'Data');
    expect(file1Sheet).toBeDefined(); // 'Data' sheet exists
    // no 'Data (Sheet 2)' within file 1
    const file1AllSheets = await readBuffer(entries[0].data);
    expect(file1AllSheets).toHaveLength(1);
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
    expect(s1.rows).toHaveLength(5); // 3 data rows + empty row + next-file note
    expect(s1.rows[0][0]).toBe(1);
    expect(s1.rows[2][0]).toBe(3);

    const [s2] = await readBuffer(entries[1].data, 'Sheet1');
    expect(s2.rows).toHaveLength(3); // prev-file note + 2 data rows
    expect(s2.rows[1][0]).toBe(4); // data starts at index 1 (after note)
    expect(s2.rows[2][0]).toBe(5);

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
    expect(result.files[0].file).toMatch(/data_1\.xlsx$/);
    expect(result.files[1].file).toMatch(/data_2\.xlsx$/);

    fs.rmSync(tmpDir, { recursive: true });
  });
});

// ── asZip() without .file() ──────────────────────────────────────────────────────

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
    // Verify it can be parsed as XLSX (not as a ZIP container):
    const [{ rows: parsed }] = await readBuffer(buffer);
    expect(parsed).toHaveLength(2);
  });
});
