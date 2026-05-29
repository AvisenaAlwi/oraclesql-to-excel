import { createRequire } from 'module';

const require = createRequire(import.meta.url);
const mod = require('./dist/index.js');

export const OracleSqlToExcel = mod.OracleSqlToExcel;
export const OracleSqlToCsv   = mod.OracleSqlToCsv;

export default OracleSqlToExcel;
