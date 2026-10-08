/* load the engine + readers in Node */
const path = require('path');
const PUBLIC = path.join(__dirname, '..', '..', '..', 'public');
require(path.join(PUBLIC, 'common', 'xml.js'));
require(path.join(PUBLIC, 'common', 'opc.js'));
require(path.join(PUBLIC, 'common', 'opc-order.js'));
const COMMON = new Set(['core', 'zip', 'sha', 'crypto', 'numfmt']);   // public/common/, shared by the suite
for (const f of ['core', 'xml', 'zip', 'sha', 'crypto', 'numfmt', 'formula', 'model', 'calc', 'fn-core', 'fn-lookup', 'fn-stat', 'fn-fin', 'fn-eng', 'preserve', 'xlsx-read', 'xlsx-write', 'csv', 'xmlss']) {
  require(path.join(PUBLIC, COMMON.has(f) ? 'common' : path.join('ledger', 'js'), f + '.js'));
}
module.exports = globalThis.L;
