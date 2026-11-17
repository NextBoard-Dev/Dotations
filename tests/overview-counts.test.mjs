import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

test('affichage bureau: quatre colonnes uniquement dans la vue generale',()=>{
 const css=fs.readFileSync('style.css','utf8');
 assert.match(css,/@media \(min-width: 981px\)\s*\{\s*body\[data-page="overview"\] \.kpi-grid\s*\{\s*grid-template-columns: repeat\(4, minmax\(0, 1fr\)\);/);
 const html=fs.readFileSync('index.html','utf8');
 for(const id of ['kpi-personnes-en-poste','kpi-effets-confies','kpi-effets-historique','kpi-effets-non-rendus']) {
   assert.equal(html.split(`id="${id}"`).length-1,1);
 }
});

test('rendu reel des compteurs: donnees, filtre, cache et liste vide',()=>{
 const source=fs.readFileSync('app.js','utf8');
 const render=source.slice(source.indexOf('function renderOverview('),source.indexOf('function buildOverviewRows('));
 const nodes=Object.fromEntries(['kpi-personnes-en-poste','kpi-effets-confies','kpi-effets-historique','kpi-effets-non-rendus','overview-table-body'].map(id=>[id,{value:null}]));
 let filter=()=>true;
 const ctx={
   document:{getElementById:id=>nodes[id]||null},state:{listRenderCache:{},tableSorts:{}},
   sortPersonsForOverview:p=>p,
   getEffectsForActiveFilters:p=>p.effects.filter(filter),
   getCurrentAssignedEffectsForActiveFilters:p=>p.effects.filter(filter).filter(e=>['ACTIF','NON RENDU'].includes(e.status)),
   getEffectReplacementCost:()=>0,getEffectStatus:(p,e)=>e.status,
   getArrivalComplementMovementMap:()=>new Map(),normalizeText:s=>s,
   getDossierStatus:p=>p.status,
   setKpiCountAnimated:(node,value)=>{node.value=value;},
   renderEffectsChart:()=>{},buildOverviewRows:()=>'',renderTableRowsProgressively:()=>{},buildEmptyTableRow:()=>'',bindPersonRowActions:()=>{},
 };
 vm.createContext(ctx);vm.runInContext(render,ctx);
 const person={id:'TEST',status:'EN POSTE',effects:['ACTIF','NON RENDU','RESTITUE','CLOTURE','PERDU','HS','VOL','DETRUIT'].map(status=>({status}))};
 const values=()=>['kpi-personnes-en-poste','kpi-effets-confies','kpi-effets-historique','kpi-effets-non-rendus'].map(id=>nodes[id].value);
 ctx.renderOverview([person]);assert.deepEqual(values(),[1,2,8,1]);
 assert.equal(ctx.renderOverview([person]),false);assert.deepEqual(values(),[1,2,8,1]);
 filter=e=>e.status==='RESTITUE';ctx.renderOverview([person]);assert.deepEqual(values(),[1,0,1,0]);
 filter=()=>true;ctx.renderOverview([person]);assert.deepEqual(values(),[1,2,8,1]);
 ctx.renderOverview([]);assert.deepEqual(values(),[0,0,0,0]);
});
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
