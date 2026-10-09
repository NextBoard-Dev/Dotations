import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import vm from "node:vm";

const source = fs.readFileSync("app.js", "utf8");
for (const docType of ["arrival", "exit"]) for (const signer of ["personnel", "representant"]) {
  test(`cycle repete effacement signature relecture: ${docType}/${signer}`, () => {
    const ctx = vm.createContext({
      normalizeText: v => String(v).toUpperCase(),
      normalizeMobileSignatureSigner: v => String(v).toLowerCase(),
      DEFAULT_SUPABASE_SIGNATURES_BUCKET: "signatures",
    });
    for (const name of ["getSupabaseSignatureRowField", "mergeSupabaseMobileSignatureRows"]) {
      const start = source.indexOf(`function ${name}(`);
      assert.ok(start >= 0, name);
      const end = source.indexOf("\n}", start);
      assert.ok(end > start, name);
      vm.runInContext(source.slice(start, end + 2), ctx);
    }
    const other = signer === "personnel" ? "representant" : "personnel";
    const untouched = {image:"other-signer",validatedAt:"2026-10-09T07:00:00Z"};
    let data = {personnes:[{id:"SIMULATION",signatures:{[docType]:{
      [signer]:{image:"",validatedAt:""},[other]:untouched
    }}}],demandesSignatureMobile:[]};
    const history = [];
    for (let cycle = 0; cycle < 3; cycle++) {
      const clearedAt = new Date(Date.UTC(2026,9,9,8,cycle*2)).toISOString();
      data.personnes[0].signatures[docType][signer] = {image:"",validatedAt:"",clearedAt};
      data = JSON.parse(JSON.stringify(data));
      ctx.mergeSupabaseMobileSignatureRows(data, history, "SIMULATION", docType);
      assert.equal(data.personnes[0].signatures[docType][signer].image,"");
      const row = {person_id:"SIMULATION",doc_type:docType,signer,
        signature_data:`new-${cycle}`,signed_at:new Date(Date.UTC(2026,9,9,8,cycle*2+1)).toISOString()};
      history.push(row);
      ctx.mergeSupabaseMobileSignatureRows(data, [row], "SIMULATION", docType);
      data = JSON.parse(JSON.stringify(data));
      ctx.mergeSupabaseMobileSignatureRows(data, [...history].reverse(), "SIMULATION", docType);
      ctx.mergeSupabaseMobileSignatureRows(data, history, "SIMULATION", docType);
      assert.equal(data.personnes[0].signatures[docType][signer].image,row.signature_data);
      assert.equal(data.personnes[0].signatures[docType][signer].validatedAt,row.signed_at);
      assert.deepEqual(data.personnes[0].signatures[docType][other],untouched);
    }
  });
}

