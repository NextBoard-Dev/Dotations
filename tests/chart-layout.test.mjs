import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { chromium } from 'playwright';

test('chart buttons retain label, proportional bars and right-aligned count', async()=>{
 const browser=await chromium.launch({channel:process.env.CI ? undefined : 'chrome',headless:true});
 try{
 const page=await browser.newPage({viewport:{width:1650,height:420}});
 const css=fs.readFileSync('style.css','utf8');
 const theme=(fs.existsSync('nextboard-theme.css') ? fs.readFileSync('nextboard-theme.css','utf8') : '');
 await page.setContent(`<style>${css}\n${theme}</style><body class="nb-theme-enabled"><div class="effects-chart"><button type="button" class="effects-chart__row" aria-pressed="false"><span class="effects-chart__label">CLE</span><span class="effects-chart__track"><span class="effects-chart__bar-group" style="width:80%"><span class="effects-chart__segment effects-chart__segment--actif" style="width:100%"></span></span></span><strong class="effects-chart__value">46</strong></button><button type="button" class="effects-chart__row" aria-pressed="true"><span class="effects-chart__label">TELECOMMANDE URMET</span><span class="effects-chart__track"><span class="effects-chart__bar-group" style="width:50%"><span class="effects-chart__segment effects-chart__segment--restitue" style="width:100%"></span></span></span><strong class="effects-chart__value">30</strong></button></div></body>`);
 for(const dark of [false,true]){
 await page.evaluate(d=>document.body.classList.toggle('theme-dark',d),dark);
 const sizes=await page.locator('.effects-chart__row').first().evaluate(row=>{const [label,track,count]=row.children;return {display:getComputedStyle(row).display,label:label.getBoundingClientRect().x,track:track.getBoundingClientRect().width,count:count.getBoundingClientRect().x,bar:track.firstElementChild.getBoundingClientRect().width};});
 assert.equal(sizes.display,'grid');assert.ok(sizes.track>1000);assert.ok(sizes.label<200);assert.ok(sizes.count>1500);assert.ok(Math.abs(sizes.bar/sizes.track-.8)<.01);
 }
 await page.evaluate(()=>document.body.classList.remove('theme-dark'));
 await page.evaluate(()=>document.getAnimations().forEach(animation=>animation.finish()));

 }finally{await browser.close();}
});
