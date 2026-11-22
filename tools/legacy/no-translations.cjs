// The task forbids access to translation catalogues. Exclude that directory from
// broad source-tree test enumeration, and fail any accidental direct access.
const fs = require('node:fs'), path = require('node:path');
const forbidden = path.resolve('public/common/i18n');
const blocked = file => typeof file === 'string' && (path.resolve(file) === forbidden || path.resolve(file).startsWith(forbidden + path.sep));
const list = fs.readdirSync, read = fs.readFileSync;
fs.readdirSync = function (dir, ...args) { if (blocked(dir)) return []; return list.call(this, dir, ...args); };
fs.readFileSync = function (file, ...args) { if (blocked(file)) throw new Error('Translation catalog access is excluded by I24'); return read.call(this, file, ...args); };
