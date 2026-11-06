import fs from 'node:fs';
import vm from 'node:vm';
import path from 'node:path';
export const root = path.resolve(import.meta.dirname, '../..');
export function load() {
  const common = path.join(root, 'public/common');
  const files = ['zip.js', 'crypto.js', 'xml.js', 'opc.js'];
  for (const file of files) vm.runInThisContext(fs.readFileSync(path.join(common, file), 'utf8'), { filename: file });
  const source = fs.readFileSync(path.join(common, 'legacy/import.js'), 'utf8');
  const literal = /const FILES\s*=\s*(\[[\s\S]*?\]);/.exec(source)?.[1];
  if (!literal) throw new Error('The browser legacy converter list is missing');
  const converters = JSON.parse(literal.replace(/'/g, '"'));
  for (const file of converters) {
    if (!/^[a-z0-9-]+$/.test(file)) throw new Error('Invalid converter file name');
    const name = path.join(common, 'legacy', file + '.js');
    vm.runInThisContext(fs.readFileSync(name, 'utf8'), { filename: name });
  }
  return globalThis.L.legacy;
}
