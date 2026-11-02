import test from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';

test('Office HTTP protocol, serialized jobs, evidence, and Windows cross-build', () => {
  const run = (args, cwd, env = {}) => {
    const r = spawnSync('go', args, { cwd: new URL(cwd, import.meta.url), encoding: 'utf8', env: { ...process.env, ...env } });
    assert.equal(r.status, 0, r.stdout + r.stderr);
  };
  run(['test', '-race', './...'], './win_test_server/');
  fs.mkdirSync(new URL('../tmp/win-test-server/', import.meta.url), { recursive: true });
  run(['build', '-o', '../../tmp/win-test-server/check.exe', '.'], '../tools/win_test_server/', { GOOS: 'windows', GOARCH: 'amd64', CGO_ENABLED: '0' });
});
