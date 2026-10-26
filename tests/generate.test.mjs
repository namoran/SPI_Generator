/* Node test: generate an SPI in JS, then have Python verify it.
 *
 * Run:
 *   python3 -m venv .venv && .venv/bin/pip install openpyxl   (once)
 *   node tests/generate.test.mjs
 *
 * Set PYTHON=.venv/bin/python if openpyxl isn't on the default python3.
 */
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const dir = path.resolve(process.argv[2] || 'tests/fixtures');
const out = path.resolve('tests/tmp');
fs.mkdirSync(out, { recursive: true });

const F = {
  rfc: path.join(dir, '5632 RFC1 VERSION 4.xlsm'),
  spiOld: path.join(dir, 'PART2-AttachmentB-SPI-138506266-SA - Copy.xlsm'),
};

const sandbox = { console, TextDecoder, TextEncoder };
vm.createContext(sandbox);
const run = (file) => vm.runInContext(fs.readFileSync(file, 'utf8'), sandbox, { filename: file });
sandbox.fflate = require('fflate');
run('docs/ppr-xlsx.js');
run('docs/data.js');
run('docs/logic.js');
run('docs/zip-repack.js');
run('docs/patch.js');

const SPI = sandbox.SPI;
const SPIPatch = sandbox.SPIPatch;
sandbox.ZipRepack = sandbox.ZipRepack;
const u8 = (p) => {
  const b = fs.readFileSync(p);
  return new Uint8Array(b.buffer, b.byteOffset, b.byteLength);
};

const rfcWb = SPI.open(u8(F.rfc), 'the RFC file');
const spiWb = SPI.open(u8(F.spiOld), 'the old SPI');
const changes = SPI.plan(
  SPI.sheetOrThrow(rfcWb, SPI.RFC_SHEET, 'the RFC file'),
  SPI.sheetOrThrow(spiWb, SPI.SOW_SHEET, 'the old SPI')
);

console.log(`plan: ${changes.length} cells to update`);
changes.slice(0, 4).forEach((c) => {
  console.log(`  ${c.address.padEnd(5)} ${c.task.padEnd(7)} ${c.payItem.padEnd(8)} ${c.oldValue} + ${c.units} = ${c.newValue}`);
});

const t0 = performance.now();
const bytes = SPIPatch.build(u8(F.spiOld), spiWb, SPI.SOW_SHEET, changes);
const buildMs = performance.now() - t0;
const actualPath = path.join(out, 'actual.xlsm');
fs.writeFileSync(actualPath, bytes);
console.log(`built ${(bytes.length / 1048576).toFixed(2)} MB in ${buildMs.toFixed(0)} ms -> ${actualPath}`);

const py = process.env.PYTHON || 'python3';
const res = spawnSync(py, ['tests/verify.py', actualPath, F.spiOld, F.rfc], { encoding: 'utf8' });
if (res.error) {
  console.error('could not run python:', res.error.message);
  process.exit(2);
}
if (res.stderr && !res.stdout) console.error(res.stderr);
const report = JSON.parse(res.stdout);
console.log(JSON.stringify({ ok: report.ok, problems: report.problems, vba_bytes: report.vba_bytes, parts: report.parts }, null, 1));
if (!report.ok) process.exit(1);
console.log('all checks passed');
