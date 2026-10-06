/* load the engine + readers in Node */
const path = require('path');
for (const f of ['core', 'xml', 'zip', 'sha', 'crypto', 'numfmt', 'formula', 'model', 'calc', 'fn-core', 'fn-lookup', 'fn-stat', 'fn-fin', 'fn-eng', 'xlsx-read', 'xlsx-write', 'csv', 'xmlss']) {
  const p = path.join(__dirname, '..', '..', '..', 'public', 'ledger', 'js', f + '.js');
  try { require(p); } catch (e) { if (e.code !== 'MODULE_NOT_FOUND') throw e; }
}
module.exports = globalThis.L;
