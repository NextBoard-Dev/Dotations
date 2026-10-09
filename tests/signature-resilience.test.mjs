import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import vm from "node:vm";

const source = fs.readFileSync("app.js", "utf8");
const start = source.indexOf("async function refreshMobileSignaturePageFromRelay()");
const end = source.indexOf("\nfunction startMobileSignaturePageRefresh()", start);
assert.ok(start >= 0 && end > start);

for (const docType of ["arrival", "exit"]) for (const signer of ["personnel", "representant"]) {
  test(`reprise reseau et isolation: ${docType}/${signer}`, async () => {
    const request = {token:"TEST-ONLY", docType, signer};
    const person = {id:"SIMULATION"};
    const row = {token:request.token,person_id:person.id,doc_type:docType,signer,signature_data:"synthetic",signed_at:"2026-10-09T12:22:00Z"};
    let fetcher = async () => { throw new Error("offline"); };
    let renders = 0;
    const ctx = {
      state:{data:{personnes:[],demandesSignatureMobile:[]}},
      document:{body:{dataset:{page:"mobile-signature"}},visibilityState:"visible"},
      console:{warn(){}},
      getCurrentMobileSignatureRequest:()=>request,
      getMobileSignatureTargetPerson:()=>person,
      isSupabaseConfigured:()=>true,
      normalizeText:v=>String(v).toUpperCase(),
      normalizeMobileSignatureSigner:v=>String(v).toLowerCase(),
      getMobileSignatureRuntimeSignature:()=>null,
      fetchSupabaseMobileSignatureRows:()=>fetcher(),
      getCurrentMobileSignatureToken:()=>request.token,
      getSupabaseSignatureRowField:(r,...keys)=>String(keys.map(k=>r[k]).find(Boolean)||""),
      setSignatureValue:(p,d,s,image)=>{p.image=image;},
      rememberMobileSignatureRuntimeSignature:()=>{},
      renderMobileSignaturePage:()=>{renders++;},
      refreshDocumentSignatureCanvases:()=>{},
    };
    vm.createContext(ctx);
    vm.runInContext(source.slice(start,end),ctx);
    await ctx.refreshMobileSignaturePageFromRelay();
    assert.equal(ctx.state.mobileSignaturePageReadInFlight,false);
    assert.equal(renders,0);
    for (const patch of [{person_id:"OTHER"},{token:"OTHER"},{doc_type:docType==="arrival"?"exit":"arrival"},{signer:signer==="personnel"?"representant":"personnel"},{signature_data:""},{signed_at:""}]) {
      fetcher=async()=>[{...row,...patch}];
      await ctx.refreshMobileSignaturePageFromRelay();
      assert.equal(renders,0);
    }
    let release;
    fetcher=()=>new Promise(resolve=>{release=resolve;});
    const pending=ctx.refreshMobileSignaturePageFromRelay();
    await ctx.refreshMobileSignaturePageFromRelay();
    release([row]);
    await pending;
    assert.equal(renders,1);
    assert.equal(person.image,"synthetic");
    assert.equal(request.status,"SIGNEE");
    assert.equal(ctx.state.mobileSignaturePageReadInFlight,false);
  });
}

