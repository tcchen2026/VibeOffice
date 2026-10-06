/* Run recalc1.js over the corpus with a small process pool. Usage: node recalc-all.js <dir> <out.jsonl> [filter] */
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
const dir = process.argv[2], outFile = process.argv[3], filt = process.argv[4] ? new RegExp(process.argv[4]) : null;
const files = fs.readdirSync(dir).filter((f) => /\.xls[xm]$/i.test(f) && (!filt || filt.test(f))).map((f) => path.join(dir, f));
const out = fs.createWriteStream(outFile);
let i = 0, active = 0, done = 0;
const POOL = 2;
function next() {
  while (active < POOL && i < files.length) {
    const f = files[i++];
    active++;
    const p = spawn('node', ['--max-old-space-size=3000', path.join(__dirname, 'recalc1.js'), f]);
    let buf = '';
    const timer = setTimeout(() => { p.kill('SIGKILL'); }, 90000);
    p.stdout.on('data', (d) => (buf += d));
    p.stderr.on('data', () => {});
    p.on('close', (code) => {
      clearTimeout(timer);
      const line = buf.trim().split('\n').pop();
      if (line && line[0] === '{') out.write(line + '\n'); else out.write(JSON.stringify({ file: path.basename(f), crash: code == null ? 'timeout' : 'exit ' + code }) + '\n');
      active--; done++;
      if (done % 100 === 0) process.stderr.write(done + '/' + files.length + '\n');
      if (done === files.length) out.end(); else next();
    });
  }
}
next();
