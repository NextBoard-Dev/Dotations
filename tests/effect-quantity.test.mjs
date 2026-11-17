import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';
const source = readFileSync(new URL('../app.js', import.meta.url), 'utf8');
const context = vm.createContext({});
for (const name of ['getNextId', 'appendEffectUnits']) {
  const start = source.indexOf(`function ${name}(`);
  const end = source.indexOf('\n}', start) + 2;
  assert(start >= 0 && end > start);
  vm.runInContext(source.slice(start, end), context);
}
test('quantite: chaque unite possede son identifiant et sa date de remise', () => {
  const effects = [{ id: 'E0004', dateRemise: '2026-09-09' }];
  context.appendEffectUnits(effects, {id:'ignored',typeEffet:'CLE',dateRemise:'2026-11-17'}, 3);
  assert.deepEqual(effects.map(e => e.id), ['E0004','E0005','E0006','E0007']);
  assert.equal(effects[0].dateRemise, '2026-09-09');
  assert(effects.slice(1).every(e => e.quantite === 1 && e.dateRemise === '2026-11-17'));
  effects[1].dateRetour = '2026-11-18';
  assert.equal(effects[2].dateRetour, undefined);
});
test('quantite: refus des valeurs invalides sans ajout partiel', () => {
  for (const quantity of [0,-1,1.5,1000,NaN,Infinity]) {
    const effects=[];
    assert.throws(() => context.appendEffectUnits(effects, {}, quantity));
    assert.equal(effects.length,0);
  }
});
