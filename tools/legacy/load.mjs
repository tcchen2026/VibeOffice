import fs from 'node:fs';
import vm from 'node:vm';
import path from 'node:path';
export const root = path.resolve(import.meta.dirname, '../..');
export function load() {
  const common = path.join(root, 'public/common');
  const files = ['zip.js', 'crypto.js', 'legacy/base.js', 'legacy/codepages.js', 'legacy/properties.js', 'legacy/xls-read.js', 'legacy/doc-read.js', 'legacy/ppt-read.js', 'legacy/inspect.js'];
  for (const file of files) vm.runInThisContext(fs.readFileSync(path.join(common, file), 'utf8'), { filename: file });
  for (const file of ['package', 'property-parts', 'images', 'officeart', 'native-geometry', 'shape-path', 'shapes', 'xls-write', 'doc-write', 'ppt-write', 'encryption', 'macros']) {
    const name = path.join(common, 'legacy', file + '.js');
    if (fs.existsSync(name)) vm.runInThisContext(fs.readFileSync(name, 'utf8'), { filename: name });
  }
  return globalThis.L.legacy;
}
