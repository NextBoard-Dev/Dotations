import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {chromium} from 'playwright';
test('overview headers toggle sorting by mouse and keyboard, including numeric costs',async()=>{
 const source=fs.readFileSync('app.js','utf8');
 const extract=name=>{const start=source.indexOf(`function ${name}(`);const end=source.indexOf('\nfunction ',start+1);return source.slice(start,end);};
 const functions=['getEffectTableSort','setEffectTableSort','getOverviewSortValue','sortPersonsForOverview','compareEffectValues','updateSortableHeaders','bindEffectTableSorting'].map(extract).join('\n');
 const headers=fs.readFileSync('index.html','utf8').match(/<th data-sort-table="overviewPersons"[^>]*>.*?<\/th>/g).join('');
 const browser=await chromium.launch({channel:process.env.CI ? undefined : 'chrome',headless:true});
 try{const page=await browser.newPage();await page.setContent(`<table><thead><tr>${headers}<th>ACTION</th></tr></thead><tbody></tbody></table>`);
 await page.addScriptTag({content:`const state={tableSorts:{}}; const DEFAULT_TABLE_SORTS={overviewPersons:{key:'nom',dir:'asc'}};
 const normalizeAmount=Number;const compareTextValues=(a,b)=>String(a).localeCompare(String(b),'fr');
 const getCurrentAssignedEffectsForActiveFilters=p=>p.effects;const getEffectsForActiveFilters=p=>p.effects;const getArrivalComplementMovementMap=()=>new Map();const getEffectReplacementCost=(p,e)=>e.cost;const getEffectStatus=(p,e)=>e.status;const getPersonSiteLabel=p=>p.site;const getDossierStatus=p=>p.status;
 const updateAllFilterResetHighlights=()=>{};const schedulePageRender=()=>updateSortableHeaders('overviewPersons');
 ${functions}
 bindEffectTableSorting('overviewPersons');updateSortableHeaders('overviewPersons');`});
 const keys=await page.locator('[data-sort-key]').evaluateAll(h=>h.map(e=>e.dataset.sortKey));assert.equal(keys.length,13);
 for(const key of keys){const h=page.locator(`[data-sort-key="${key}"]`);await h.click();const first=await h.getAttribute('aria-sort');await h.click();assert.notEqual(await h.getAttribute('aria-sort'),first);await h.press('Enter');assert.equal(await h.getAttribute('aria-sort'),first);}
 const sorted=await page.evaluate(()=>{state.tableSorts.overviewPersons={key:'couts',dir:'asc'};const p=[{id:'A',effects:[{cost:100}]},{id:'B',effects:[{cost:9}]},{id:'C',effects:[{cost:20}]}];const asc=sortPersonsForOverview(p).map(x=>x.id);state.tableSorts.overviewPersons.dir='desc';return [asc,sortPersonsForOverview(p).map(x=>x.id)];});assert.deepEqual(sorted,[['B','C','A'],['A','C','B']]);
 }finally{await browser.close();}
});
