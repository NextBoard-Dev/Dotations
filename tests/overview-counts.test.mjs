import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
test('compteurs distincts: affectation courante et historique filtre',()=>{
 const source=fs.readFileSync('app.js','utf8');
 const render=source.slice(source.indexOf('function renderOverview('),source.indexOf('function buildOverviewRows('));
 assert.match(render,/totalEffectsCount \+= assignedEffects.length/);
 assert.match(render,/historicalEffectsCount \+= 1/);
 assert.doesNotMatch(render,/totalEffectsCount \+= 1/);
 assert.match(render,/setKpiCountAnimated\(historicalEffectsNode, historicalEffectsCount\)/);
 const html=fs.readFileSync('index.html','utf8');
 assert.equal((html.match(/id="kpi-effets-historique"/g)||[]).length,1);
 assert.match(html,/TOTAL HISTORIQUE/);
});
