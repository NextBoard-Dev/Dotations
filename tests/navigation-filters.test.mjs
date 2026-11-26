import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const source=fs.readFileSync('app.js','utf8');
const code=source.slice(source.indexOf('function getStoredNavigationContext('),source.indexOf('function getCurrentPersonId('));
function run({referrer='',type='navigate',stored=true,query='' }={}){
 const defaults={search:'',site:'',typeEffet:''};let raw=stored?JSON.stringify({personId:'P1',urgentMode:true,filters:{search:'MARTIN',site:'CDM',typeEffet:'CLE'}}):null;
 const ctx={URL,URLSearchParams,NAVIGATION_CONTEXT_KEY:'nav',DEFAULT_FILTERS:defaults,state:{},document:{referrer,body:{dataset:{page:'overview'}}},window:{location:{href:'https://dashboard.example.test/index.html'+query,search:query},performance:{getEntriesByType:()=>[{type}]},sessionStorage:{getItem:()=>raw,setItem:(k,v)=>{raw=v}},history:{replaceState(){}}},setCurrentPersonId:id=>{ctx.selected=id;}};
 vm.createContext(ctx);vm.runInContext(code,ctx);ctx.restoreNavigationContext();return {ctx,saved:JSON.parse(raw)};
}
test('new dashboard opening clears all saved filters and selection',()=>{for(const referrer of ['', 'https://home.example.test/', 'https://outside.test/']){const {ctx,saved}=run({referrer});assert.equal(ctx.state.filters.typeEffet,'');assert.equal(ctx.state.filters.search,'');assert.equal(ctx.state.urgentMode,false);assert.equal(saved.personId,'');}});
test('internal navigation and history preserve filters and selected person',()=>{for(const options of [{referrer:'https://dashboard.example.test/fiche-personne.html'},{type:'back_forward'}]){const {ctx}=run(options);assert.equal(ctx.state.filters.search,'MARTIN');assert.equal(ctx.state.filters.typeEffet,'CLE');assert.equal(ctx.state.filters.site,'CDM');assert.equal(ctx.selected,'P1');}});
test('fresh deep link keeps explicit person while resetting old filters',()=>{const {ctx,saved}=run({query:'?personId=P2'});assert.equal(saved.personId,'P2');assert.equal(ctx.state.filters.typeEffet,'');});

test('F5 reload resets filters even with an internal referrer',()=>{const {ctx,saved}=run({type:'reload',referrer:'https://dashboard.example.test/fiche-personne.html'});assert.equal(ctx.state.filters.typeEffet,'');assert.equal(ctx.state.filters.search,'');assert.equal(saved.personId,'');assert.equal(ctx.state.urgentMode,false);});
