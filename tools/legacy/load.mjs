import fs from 'node:fs';
import vm from 'node:vm';
import path from 'node:path';
export const root = path.resolve(import.meta.dirname, '../..');
export function load() {
  const common = path.join(root, 'public/common');
  const files = ['zip.js', 'crypto.js', 'xml.js', 'opc.js', 'legacy/base.js', 'legacy/codepages.js', 'legacy/properties.js', 'legacy/xls-read.js', 'legacy/doc-read.js', 'legacy/ppt-read.js', 'legacy/inspect.js'];
  for (const file of files) vm.runInThisContext(fs.readFileSync(path.join(common, file), 'utf8'), { filename: file });
  for (const file of ['package', 'property-parts', 'images', 'officeart', 'native-geometry', 'shape-path', 'shapes', 'xml-check', 'xls-style-ext', 'xls-style', 'xls-sheet', 'xls-functions', 'xls-formula', 'xls-formula-book', 'xls-write', 'doc-sprm', 'doc-structure', 'doc-style', 'doc-lists', 'doc-xml', 'doc-settings', 'doc-text', 'doc-write', 'ppt-structure', 'ppt-theme', 'ppt-text-read', 'ppt-text', 'ppt-text-design', 'ppt-text-write', 'ppt-write', 'encryption', 'macros']) {
    const name = path.join(common, 'legacy', file + '.js');
    if (fs.existsSync(name)) vm.runInThisContext(fs.readFileSync(name, 'utf8'), { filename: name });
  }
  return globalThis.L.legacy;
}
