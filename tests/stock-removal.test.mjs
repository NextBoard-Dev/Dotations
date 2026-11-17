import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const read = name => readFileSync(new URL('../' + name, import.meta.url), 'utf8');

test('stocks retires: aucun calcul, mouvement automatique ou controle actif', () => {
  assert.doesNotMatch(read('app.js'), /stocksEffetsManuels|function\s+\w*Stock\w*\s*\(|stockTableFilters|stockHighlightKey/);
  assert.doesNotMatch(read('bases-reference.html'), /id="stock-/);
  assert.doesNotMatch(read('style.css'), /\.stock-|\.kpi-grid--stock/);
});

test('retrait des stocks: catalogue, tarifs et effets confies conserves', () => {
  const app = read('app.js');
  assert.match(app, /function renderReferenceEffectsTable\(/);
  assert.match(app, /function renderReplacementCostsTable\(/);
  assert.match(app, /person\.effetsConfies\.push\(effect\)/);
  assert.match(read('index.html'), /id="kpi-effets-confies"/);
  assert.match(read('index.html'), /id="kpi-effets-historique"/);
});
