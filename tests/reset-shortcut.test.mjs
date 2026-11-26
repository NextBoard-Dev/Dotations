import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {chromium} from 'playwright';
test('Ctrl Alt R uses native filter reset from a focused input',async()=>{
 const source=fs.readFileSync('app.js','utf8');const start=source.indexOf('function bindGlobalShortcuts()');const code=source.slice(start,source.indexOf('function bindHistoryNavigation()',start));
 const browser=await chromium.launch({channel:process.env.CI ? undefined : 'chrome',headless:true});
 try{const page=await browser.newPage();await page.setContent('<form class="js-filter-form"><input name="search"><select name="site"><option value="">TOUS</option><option>CDM</option></select></form>');
 await page.addScriptTag({content:`const state={};let resets=0;function undoLastChange(){};document.querySelector('form').addEventListener('reset',()=>resets++);${code};bindGlobalShortcuts();bindGlobalShortcuts();`});
 await page.locator('input').fill('KOUAME');await page.locator('select').selectOption('CDM');await page.locator('input').press('Control+Alt+r');assert.equal(await page.locator('input').inputValue(),'');assert.equal(await page.locator('select').inputValue(),'');assert.equal(await page.evaluate(()=>resets),1);
 await page.locator('input').fill('test');await page.locator('input').press('r');assert.equal(await page.evaluate(()=>resets),1);
 }finally{await browser.close();}
});
