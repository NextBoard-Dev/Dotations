import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

test('les deux controles locaux ont une copie versionnee non vide', () => {
  for (const name of ['CHECK_LOCAL.ps1', 'CHECK_PDF_INTEGRITY.ps1']) {
    const source = fs.readFileSync(`scripts/local-guards/${name}`, 'utf8');
    assert.ok(source.length > 500);
    assert.match(source, /exit 1/);
    assert.match(source, /local_state\.json/);
  }
});

test('les controles conservent leurs verifications bloquantes', () => {
  const local = fs.readFileSync('scripts/local-guards/CHECK_LOCAL.ps1', 'utf8');
  const pdf = fs.readFileSync('scripts/local-guards/CHECK_PDF_INTEGRITY.ps1', 'utf8');
  assert.match(local, /node --check/);
  assert.match(local, /ContainsKey/);
  assert.match(pdf, /%PDF-/);
  assert.match(pdf, /%%EOF/);
  assert.match(pdf, /StartsWith/);
  assert.match(pdf, /Test-Path/);
});
