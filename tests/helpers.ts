import ExcelJS from 'exceljs';

export interface MockMeta {
  name: string;
  dbTypeName?: string;
}

/**
 * Build a mock Oracle connection that streams `rows` via resultSet.getRows().
 * Each call to getRows() returns up to `fetchSize` rows, then [] to signal EOF.
 */
export function createStreamConn(
  rows: Record<string, unknown>[],
  meta: MockMeta[] = []
) {
  let offset = 0;

  const resultSet = {
    async getRows(count: number) {
      const batch = rows.slice(offset, offset + count);
      offset += batch.length;
      return batch;
    },
    async close() {},
  };

  return {
    async execute(_sql: string, _params?: unknown, _opts?: unknown) {
      return { metaData: meta, resultSet };
    },
    async close() {},
  };
}

/**
 * Build a mock Oracle connection for SELECT COUNT(*) queries.
 * Returns `{ TOTAL: total }` from rows[0].
 */
export function createCountConn(total: number) {
  return {
    async execute(_sql: string, _params?: unknown, _opts?: unknown) {
      return { rows: [{ TOTAL: total }] };
    },
    async close() {},
  };
}

/**
 * Parse an XLSX buffer with ExcelJS and return rows as plain objects keyed by header.
 */
export async function readBuffer(
  buffer: Buffer,
  sheetName?: string
): Promise<{ headers: string[]; rows: unknown[][] }[]> {
  const wb = new ExcelJS.Workbook();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  await wb.xlsx.load(buffer as any);

  const results: { headers: string[]; rows: unknown[][] }[] = [];

  const sheets = sheetName
    ? [wb.getWorksheet(sheetName)].filter(Boolean)
    : wb.worksheets;

  for (const ws of sheets) {
    if (!ws) continue;
    const allRows = ws.getSheetValues() as unknown[][];
    // getSheetValues() returns 1-indexed array (index 0 is undefined)
    const dataRows = allRows.filter(Boolean);
    const headers = (dataRows[0] as unknown[] | undefined)?.filter(
      (v): v is string => typeof v === 'string'
    ) ?? [];
    const rows = dataRows.slice(1).map((r) => (r as unknown[]).slice(1));
    results.push({ headers, rows });
  }

  return results;
}
