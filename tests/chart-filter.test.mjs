import fs from 'node:fs';
import test from 'node:test';
import assert from 'node:assert/strict';
import {chromium} from 'playwright';
test('real chart segments combine type and status without changing other filters or layout',async()=>{
 const source=fs.readFileSync('app.js','utf8');
 const extract=name=>{const start=source.indexOf(`function ${name}(`);return source.slice(start,source.indexOf('\nfunction ',start+1));};
 const code=['renderEffectsChart','getEffectChartCategory','effectMatchesActiveObjectFilters','getEffectsForActiveFilters'].map(extract).join('\n');
 const browser=await chromium.launch({channel:process.env.CI ? undefined : 'chrome',headless:true});
 try{const page=await browser.newPage({viewport:{width:1650,height:500}});
 const statuses=['ACTIF','RESTITUE','NON RENDU','VOL','PERDU','HS','CLOTURE','DETRUIT'];
 await page.setContent(`<style>${fs.readFileSync('style.css','utf8')}\n${(fs.existsSync('nextboard-theme.css') ? fs.readFileSync('nextboard-theme.css','utf8') : '')}</style><body class="nb-theme-enabled"><form class="js-filter-form"><select name="typeEffet"><option value="">TOUS</option><option>VENTILATEUR</option><option>CLE</option></select><select name="statutObjet"><option value="">TOUS</option>${statuses.map(s=>`<option>${s}</option>`).join('')}</select></form><div id="chart"></div></body>`);
 await page.addScriptTag({content:`
 const state={filters:{site:'CDM',typeEffet:'',statutObjet:''},data:{listes:{typesEffets:['VENTILATEUR','CLE']}}}; const DEFAULT_FILTERS={};
 const statuses=${JSON.stringify(statuses)};
 const persons=[{effetsConfies:statuses.map(status=>({typeEffet:'VENTILATEUR',status})).concat([{typeEffet:'CLE',status:'ACTIF'}])}];
 const getAllEffects=ps=>ps.flatMap(person=>person.effetsConfies.map(effect=>({person,effect})));
 const effectMatchesSiteFilter=()=>true;const normalizeText=s=>s;const getEffectStatus=(p,e)=>e.status;
 const getReplacementCostValue=()=>0;const getEffectReplacementCost=()=>0;const formatAmountWithEuro=String;const escapeHtml=s=>s;
 ${code}
 document.querySelector('form').addEventListener('input',()=>{state.filters.typeEffet=document.querySelector('[name=typeEffet]').value;state.filters.statutObjet=document.querySelector('[name=statutObjet]').value;renderEffectsChart('chart',persons);});
 renderEffectsChart('chart',persons);
 `});
 const row=page.locator('[data-effect-type="VENTILATEUR"]');
 const initial=await row.evaluate(r=>({display:getComputedStyle(r).display,width:r.querySelector('.effects-chart__track').getBoundingClientRect().width}));assert.equal(initial.display,'grid');assert.ok(initial.width>1000);
 assert.equal(await page.locator('button button').count(),0);
 for(const status of statuses){
  await row.locator(`[data-effect-status="${status}"]`).click();
  assert.equal(await page.locator('[name=typeEffet]').inputValue(),'VENTILATEUR');assert.equal(await page.locator('[name=statutObjet]').inputValue(),status);
  const appearance=await row.evaluate(r=>{const label=getComputedStyle(r.querySelector('.effects-chart__label'));const segment=getComputedStyle(r.querySelector('[aria-pressed="true"]'));return {outline:getComputedStyle(r).outlineStyle,weight:label.fontWeight,shadow:segment.boxShadow,filter:segment.filter};});assert.equal(appearance.outline,'none');assert.equal(appearance.weight,'400');assert.equal(appearance.shadow,'none');assert.equal(appearance.filter,'none');
  const result=await page.evaluate(()=>({effects:getEffectsForActiveFilters(persons[0]).map(e=>e.status),site:state.filters.site}));assert.deepEqual(result,{effects:[status],site:'CDM'});
  assert.equal(await row.locator('.effects-chart__value').textContent(),'1');assert.equal(await row.locator('[data-effect-status]').getAttribute('aria-pressed'),'true');
  await row.locator(`[data-effect-status="${status}"]`).click();assert.equal(await page.locator('[name=statutObjet]').inputValue(),'');assert.equal(await page.locator('[name=typeEffet]').inputValue(),'VENTILATEUR');
 }
 await row.locator('[data-effect-status="ACTIF"]').press('Enter');assert.equal(await page.locator('[name=statutObjet]').inputValue(),'ACTIF');
 await row.locator('.effects-chart__label').click();assert.equal(await page.locator('[name=statutObjet]').inputValue(),'');assert.equal(await page.locator('[name=typeEffet]').inputValue(),'VENTILATEUR');
 await row.locator('.effects-chart__label').click();assert.equal(await page.locator('[name=typeEffet]').inputValue(),'');
 await page.evaluate(()=>document.getAnimations().forEach(a=>a.finish()));

 }finally{await browser.close();}
});
