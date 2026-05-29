import { OracleSqlToExcel } from '../src/index';
import { createStreamConn } from './helpers';

const conn = () => Promise.resolve(createStreamConn([]));

describe('OracleSqlToExcelBuilder — validation', () => {
  it('throws when no connectionFactory is set', async () => {
    const result = await OracleSqlToExcel()
      .sheet('Sheet1', (s) => s.sql('SELECT 1 FROM DUAL'))
      .toBuffer();
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/connection factory/i);
  });

  it('throws when no sheets defined', async () => {
    const result = await OracleSqlToExcel()
      .connectionFactory(conn)
      .toBuffer();
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/no sheets/i);
  });

  it('throws when sheet has no SQL', async () => {
    const result = await OracleSqlToExcel()
      .connectionFactory(conn)
      .sheet('Sheet1', (_s) => {})
      .toBuffer();
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/no sql query/i);
  });

  it('throws synchronously when sheet name exceeds 25 chars', () => {
    expect(() => {
      OracleSqlToExcel().sheet('A'.repeat(26), (_s) => {});
    }).toThrow(/exceeds 25 characters/i);
  });

  it('accepts sheet name at exactly 25 chars', () => {
    expect(() => {
      OracleSqlToExcel().sheet('A'.repeat(25), (_s) => {});
    }).not.toThrow();
  });

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

  it('backpressureThreshold default is 256 MB', () => {
    const builder = OracleSqlToExcel() as any;
    expect(builder._backpressureThreshold).toBe(256 * 1024 * 1024);
  });

  it('compress() default level is 1', () => {
    const builder = OracleSqlToExcel() as any;
    builder.compress(true);
    expect(builder._compressLevel).toBe(1);
  });

  it('compress(true, n) stores custom level', () => {
    const builder = OracleSqlToExcel() as any;
    builder.compress(true, 6);
    expect(builder._compress).toBe(true);
    expect(builder._compressLevel).toBe(6);
  });

  it('compress() throws RangeError for level below 0', () => {
    expect(() => OracleSqlToExcel().compress(true, -1)).toThrow(RangeError);
  });

  it('compress() throws RangeError for level above 9', () => {
    expect(() => OracleSqlToExcel().compress(true, 10)).toThrow(RangeError);
  });

  it('compress() throws RangeError for non-integer level', () => {
    expect(() => OracleSqlToExcel().compress(true, 1.5)).toThrow(RangeError);
  });
});
