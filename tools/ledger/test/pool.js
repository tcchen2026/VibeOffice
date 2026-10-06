/* Run a per-file test script over a directory with a small process pool.
 * Usage: node pool.js <script.js> <dir> <out.jsonl> [filterRegex] [extraArg] */
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
const [script, dir, outFile, filtS, extra] = process.argv.slice(2);
const filt = filtS && filtS !== '-' ? new RegExp(filtS) : null;
const files = fs.readdirSync(dir).filter((f) => /\.(xls[xm]|csv|txt|tsv)$/i.test(f) && (!filt || filt.test(f))).map((f) => path.join(dir, f));
const out = fs.createWriteStream(outFile);
let i = 0, active = 0, done = 0;
const POOL = 2;
function next() {
  while (active < POOL && i < files.length) {
    const f = files[i++];
    active++;
    const args = ['--max-old-space-size=3000', path.resolve(script), f];
    if (extra) args.push(extra);
    const p = spawn('node', args);
    let buf = '';
    const timer = setTimeout(() => { p.kill('SIGKILL'); }, 120000);
    p.stdout.on('data', (d) => (buf += d));
    p.stderr.on('data', () => {});
    p.on('close', (code) => {
      clearTimeout(timer);
      const line = buf.trim().split('\n').pop();
      if (line && line[0] === '{') out.write(line + '\n'); else out.write(JSON.stringify({ file: path.basename(f), crash: code == null ? 'timeout' : 'exit ' + code }) + '\n');
      active--; done++;
      if (done % 200 === 0) process.stderr.write(done + '/' + files.length + '\n');
      if (done === files.length) out.end(); else next();
    });
  }
}
next();
