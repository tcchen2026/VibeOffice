// Fresh-process timing, including sniff, for a legacy document.
// node tools/legacy/time.mjs INPUT --out RESULT.json [--passwords PASSWORD_MAP.json]
import fs from 'node:fs';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { load } from './load.mjs';

const args = process.argv.slice(2);
const option = name => {
  const index = args.indexOf(name);
  return index < 0 ? undefined : args[index + 1];
};
const file = args[0];
if (!file || file.startsWith('--')) throw Error('Provide a binary Office input file');
const bytes = fs.readFileSync(file);
const passwords = option('--passwords') ? JSON.parse(fs.readFileSync(option('--passwords'), 'utf8')) : {};
const source = path.relative('/work/corpora', path.resolve(file));
const B = load();
const before = process.memoryUsage();
const start = performance.now();
const sniff = B.sniff(bytes);
const sniffMs = performance.now() - start;
const afterSniff = process.memoryUsage();
const result = await B[sniff.format].convert(bytes, {
  extension: path.extname(file), password: passwords[source]
});
const totalMs = performance.now() - start;
const row = {
  file, sha256: createHash('sha256').update(bytes).digest('hex'), inputBytes: bytes.length,
  outputBytes: result.pkg.length, type: result.type, sniff, sniffMs,
  conversionMs: totalMs - sniffMs, totalMs, peakRSS: process.resourceUsage().maxRSS * 1024,
  before, afterSniff, after: process.memoryUsage(), stats: result.stats
};
const json = JSON.stringify(row, null, 2) + '\n';
if (option('--out')) fs.writeFileSync(option('--out'), json);
process.stdout.write(json);
