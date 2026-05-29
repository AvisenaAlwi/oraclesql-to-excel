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
