import { createRequire } from 'module';

const require = createRequire(import.meta.url);
const { OracleSqlToExcel } = require('./dist/index.js');

export { OracleSqlToExcel };
export default OracleSqlToExcel;
