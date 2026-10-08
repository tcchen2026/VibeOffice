// Independent package consistency checks; Python's strict XML parser complements the SDK validator.
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
const result = spawnSync(process.env.PYTHON || 'python3', [fileURLToPath(new URL('./package.py', import.meta.url)), 'check', ...process.argv.slice(2)], { stdio: 'inherit' });
if (result.error) console.error(result.error.message);
process.exitCode = result.status ?? 1;
