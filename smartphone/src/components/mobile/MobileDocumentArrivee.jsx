import React, { useState, useEffect } from "react";
import { db } from "@/lib/db";
import MobilePersonSearch from "./MobilePersonSearch";
import MobileSignatureCanvas from "./MobileSignatureCanvas";
import MobileEffetsList from "./MobileEffetsList";

const card = { background: "rgba(244,241,234,0.98)", border: "1px solid rgba(173,190,199,0.98)", borderRadius: 11, padding: "12px", marginBottom: 8, boxShadow: "0 4px 12px rgba(31,49,59,0.10)" };
const docField = { display: "flex", flexDirection: "column", gap: 3, marginBottom: 8 };
const docLabel = { fontSize: 9, color: "#4a6170", letterSpacing: "0.08em" };
const docValue = { padding: "7px 12px", borderRadius: 9, background: "rgba(251,250,247,0.98)", border: "1px solid rgba(152,177,190,0.9)", fontSize: 12, color: "#0f1e26", minHeight: 32, display: "flex", alignItems: "center" };
const eur = new Intl.NumberFormat("fr-FR", { style: "currency", currency: "EUR", minimumFractionDigits: 2, maximumFractionDigits: 2 });

function normalizeLabel(value) {
  return String(value || "").replace(/\s+/g, " ").trim().toUpperCase();
}

export default function MobileDocumentArrivee({ persons, effets, selectedPerson, onSelectPerson, setSaveStatus, onDataChange, representatives = [], pricingRules = [], effetTypes = [] }) {
  const [signatures, setSignatures] = useState([]);
  const [activeSection, setActiveSection] = useState("identite");
  const [representantId, setRepresentantId] = useState("");

  useEffect(() => {
    if (selectedPerson) loadSignatures();
  }, [selectedPerson]);

  useEffect(() => {
    setRepresentantId("");
  }, [selectedPerson?.id]);

  const loadSignatures = async () => {
    const sigs = await db.Signature.filter({ personId: selectedPerson.id, docType: "arrival" });
    setSignatures(sigs);
  };

  const handleSignatureSaved = async () => {
    await loadSignatures();
    if (onDataChange) await onDataChange();
  };

  const personEffets = selectedPerson ? effets.filter(e => e.personId === selectedPerson.id && e.statut === "ACTIF") : [];
  const totalValeur = personEffets.reduce((s, e) => s + (Number(e.coutRemplacement) || 0), 0);
  const costByType = new Map();
  pricingRules.forEach((rule) => {
    const type = normalizeLabel(rule.typeEffet);
    if (!type) return;
    const cause = normalizeLabel(rule.cause);
    const amount = Number(rule.montant) || 0;
    const current = costByType.get(type) || { hs: 0, perdu: 0, vol: 0, nonRendu: 0 };
    if (cause === "HS") current.hs = amount;
    if (cause === "PERTE" || cause === "PERDU") current.perdu = amount;
    if (cause === "VOL") current.vol = amount;
    if (cause === "NON RENDU") current.nonRendu = amount;
    costByType.set(type, current);
  });
  const baseTypes = Array.from(new Set(
    [...effetTypes.map((t) => normalizeLabel(t)), ...Array.from(costByType.keys())].filter(Boolean)
  ));
  const clauseRows = baseTypes.map((type) => {
    const c = costByType.get(type) || { hs: 0, perdu: 0, vol: 0, nonRendu: 0 };
    return {
      id: type,
      label: type,
      hs: c.hs,
      perdu: c.perdu,
      vol: c.vol,
      nonRendu: c.nonRendu,
    };
  });

  const getSig = (signer) => signatures.find(s => s.signer === signer);
  const personnelSig = getSig("personnel");
  const representantSig = getSig("representant");
  const isDocumentSigned = Boolean(personnelSig?.signedAt && representantSig?.signedAt);
  const documentDateIso = isDocumentSigned
    ? [personnelSig?.signedAt, representantSig?.signedAt].filter(Boolean).sort().slice(-1)[0]
    : new Date().toISOString();
  const representant = (representatives || []).find((p) => p.id === representantId) || null;
  const representantName = representant ? `${representant.nom || ""}`.trim() : "";
  const representantFunction = representant?.fonction || "";

  const fmt = (d) => d ? new Date(d).toLocaleDateString("fr-FR") : "—";

  return (
    <div style={{ padding: "12px 12px 0" }}>
      <MobilePersonSearch persons={persons} selectedPerson={selectedPerson} onSelectPerson={onSelectPerson} />

      {/* Section tabs */}
      <div style={{ display: "flex", gap: 4, marginBottom: 8, overflowX: "auto" }}>
        {[["identite", "IDENTITE"], ["effets", `EFFETS (${personEffets.length})`], ["signatures", "SIGNATURES"]].map(([id, lbl]) => (
          <button key={id} onClick={() => setActiveSection(id)} style={{ flex: "0 0 auto", padding: "7px 12px", borderRadius: 8, border: "1px solid rgba(63,97,112,0.25)", background: activeSection === id ? "rgba(63,97,112,0.22)" : "rgba(63,97,112,0.08)", color: activeSection === id ? "#213b48" : "#3f5662", fontSize: 10, fontWeight: activeSection === id ? 700 : 400, cursor: "pointer", whiteSpace: "nowrap" }}>
            {lbl}
          </button>
        ))}
      </div>

      {!selectedPerson && (
        <div style={{ ...card, textAlign: "center", color: "#3f5662", fontSize: 11 }}>SELECTIONNEZ UNE PERSONNE</div>
      )}

      {selectedPerson && activeSection === "identite" && (
        <div style={card}>
          <div style={{ marginBottom: 10, paddingBottom: 8, borderBottom: "1px solid rgba(173,190,199,0.5)" }}>
            <div style={{ fontSize: 9, color: "#556d79", letterSpacing: "0.12em" }}>DOCUMENT D'ARRIVEE</div>
            <div style={{ fontSize: 14, fontWeight: 700, color: "#14242c" }}>ATTESTATION DE REMISE DES EFFETS</div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 10px" }}>
            {[
              ["NOM", selectedPerson.nom],
              ["PRENOM", selectedPerson.prenom],
              ["FONCTION", selectedPerson.fonction],
              ["TYPE PERSONNEL", selectedPerson.typePersonnel],
              ["TYPE CONTRAT", selectedPerson.typeContrat],
              ["DATE DU DOCUMENT", fmt(documentDateIso)],
              ["DATE D'ARRIVEE", fmt(selectedPerson.dateEntree)],
              ["SORTIE PREVUE", fmt(selectedPerson.dateSortiePrevue)],
              ["SITES", (selectedPerson.sites || []).join(", ")],
            ].map(([lbl, val]) => (
              <div key={lbl} style={docField}>
                <span style={docLabel}>{lbl}</span>
                <div style={docValue}>{val || "—"}</div>
              </div>
            ))}
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginTop: 8, padding: "8px", background: "rgba(221,231,235,0.6)", borderRadius: 9 }}>
            <div>
              <div style={{ fontSize: 9, color: "#4a6170" }}>EFFETS REMIS</div>
              <div style={{ fontSize: 22, fontWeight: 700 }}>{personEffets.length}</div>
            </div>
            <div>
              <div style={{ fontSize: 9, color: "#4a6170" }}>VALEUR TOTALE</div>
              <div style={{ fontSize: 18, fontWeight: 700, color: "#9b5a2a" }}>{totalValeur.toFixed(2)} €</div>
            </div>
          </div>
        </div>
      )}

      {selectedPerson && activeSection === "effets" && (
        <div>
          <div style={{ ...card, padding: "8px 12px", marginBottom: 6 }}>
            <div style={{ fontSize: 10, color: "#3f5662", fontWeight: 600 }}>EFFETS REMIS A L'ARRIVEE (statut ACTIF)</div>
          </div>
          {personEffets.length === 0 ? (
            <div style={{ ...card, textAlign: "center", color: "#3f5662", fontSize: 11 }}>AUCUN EFFET ACTIF</div>
          ) : (
            personEffets.map(e => (
              <div key={e.id} style={{ ...card, padding: "10px 12px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 3 }}>
                  <span style={{ fontWeight: 700, fontSize: 12 }}>{e.typeEffet}</span>
                  {e.coutRemplacement && <span style={{ fontSize: 11, color: "#9b5a2a", fontWeight: 600 }}>{Number(e.coutRemplacement).toFixed(2)} €</span>}
                </div>
                <div style={{ fontSize: 11, color: "#213b48" }}>{e.designation || "—"}</div>
                {e.numeroIdentification && <div style={{ fontSize: 10, color: "#556d79" }}>N° {e.numeroIdentification}</div>}
                {e.dateRemise && <div style={{ fontSize: 10, color: "#556d79" }}>Remis le {fmt(e.dateRemise)}</div>}
              </div>
            ))
          )}
        </div>
      )}

      {selectedPerson && activeSection === "signatures" && (
        <div>
          <div style={{ ...card, padding: "10px 12px", marginBottom: 8 }}>
            <div style={{ fontSize: 10, color: "#3f5662", letterSpacing: "0.08em", marginBottom: 4, fontWeight: 600 }}>DATE DU DOCUMENT</div>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: "#14242c" }}>{fmt(documentDateIso)}</div>
              <span style={{ fontSize: 9, padding: "2px 8px", borderRadius: 99, background: isDocumentSigned ? "rgba(111,157,120,0.2)" : "rgba(217,137,106,0.2)", color: isDocumentSigned ? "#2e6a44" : "#8f4a32", border: `1px solid ${isDocumentSigned ? "rgba(111,157,120,0.3)" : "rgba(217,137,106,0.3)"}` }}>
                {isDocumentSigned ? "SIGNE" : "EN ATTENTE"}
              </span>
            </div>
          </div>
          <div style={card}>
            <div style={{ fontSize: 11, color: "#14242c", fontWeight: 700, letterSpacing: "0.08em", marginBottom: 8 }}>CLAUSES DE RESTITUTION ET DE FACTURATION</div>
            <div style={{ fontSize: 11, color: "#213b48", marginBottom: 8 }}>
              En cas de perte, vol ou détérioration imputable à une faute de l’agent, l’établissement pourra engager toute procédure administrative ou judiciaire permettant d’obtenir réparation du préjudice subi.
            </div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 520 }}>
                <thead>
                  <tr>
                    <th style={{ textAlign: "left", fontSize: 10, color: "#4a6170", padding: "4px 6px", borderBottom: "1px solid rgba(152,177,190,0.8)" }}>TYPE D'EFFET</th>
                    <th style={{ textAlign: "right", fontSize: 10, color: "#4a6170", padding: "4px 6px", borderBottom: "1px solid rgba(152,177,190,0.8)" }}>HS</th>
                    <th style={{ textAlign: "right", fontSize: 10, color: "#4a6170", padding: "4px 6px", borderBottom: "1px solid rgba(152,177,190,0.8)" }}>PERDU</th>
                    <th style={{ textAlign: "right", fontSize: 10, color: "#4a6170", padding: "4px 6px", borderBottom: "1px solid rgba(152,177,190,0.8)" }}>VOL</th>
                    <th style={{ textAlign: "right", fontSize: 10, color: "#4a6170", padding: "4px 6px", borderBottom: "1px solid rgba(152,177,190,0.8)" }}>NON RENDU</th>
                  </tr>
                </thead>
                <tbody>
                  {clauseRows.map((row) => (
                    <tr key={row.id}>
                      <td style={{ fontSize: 11, color: "#213b48", padding: "6px", borderBottom: "1px solid rgba(152,177,190,0.5)" }}>{row.label}</td>
                      <td style={{ fontSize: 11, color: "#213b48", padding: "6px", textAlign: "right", borderBottom: "1px solid rgba(152,177,190,0.5)" }}>{eur.format(row.hs)}</td>
                      <td style={{ fontSize: 11, color: "#213b48", padding: "6px", textAlign: "right", borderBottom: "1px solid rgba(152,177,190,0.5)" }}>{eur.format(row.perdu)}</td>
                      <td style={{ fontSize: 11, color: "#213b48", padding: "6px", textAlign: "right", borderBottom: "1px solid rgba(152,177,190,0.5)" }}>{eur.format(row.vol)}</td>
                      <td style={{ fontSize: 11, color: "#213b48", padding: "6px", textAlign: "right", borderBottom: "1px solid rgba(152,177,190,0.5)" }}>{eur.format(row.nonRendu)}</td>
                    </tr>
                  ))}
                  {clauseRows.length === 0 && (
                    <tr>
                      <td colSpan={5} style={{ fontSize: 11, color: "#3f5662", padding: "8px 6px", textAlign: "center" }}>AUCUN EFFET</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
          <MobileSignatureCanvas
            personId={selectedPerson.id}
            docType="arrival"
            signer="personnel"
            signerLabel="SIGNATURE DU PERSONNEL"
            existingSignature={getSig("personnel")}
            signataireName={`${selectedPerson.nom || ""} ${selectedPerson.prenom || ""}`.trim()}
            signataireFunction={selectedPerson.fonction || ""}
            onSaved={handleSignatureSaved}
          />
          <div style={{ ...card, marginBottom: 8 }}>
            <div style={{ fontSize: 10, color: "#3f5662", letterSpacing: "0.08em", marginBottom: 6, fontWeight: 600 }}>CHOIX REPRESENTANT SIGNATAIRE</div>
            <select
              value={representantId}
              onChange={(e) => setRepresentantId(e.target.value)}
              style={{ width: "100%", padding: "8px 10px", borderRadius: 9, border: "1px solid rgba(152,177,190,0.9)", background: "rgba(251,250,247,0.98)", fontSize: 11, color: "#0f1e26" }}
            >
              <option value="">SELECTIONNER UN REPRESENTANT</option>
              {(representatives || []).map((p) => (
                <option key={p.id} value={p.id}>{`${p.nom || ""}`.trim()}{p.fonction ? ` - ${p.fonction}` : ""}</option>
              ))}
            </select>
          </div>
          <MobileSignatureCanvas
            personId={selectedPerson.id}
            docType="arrival"
            signer="representant"
            signerLabel="SIGNATURE DU REPRESENTANT DE L'ETABLISSEMENT"
            existingSignature={getSig("representant")}
            signataireId={representantId}
            signataireName={representantName}
            signataireFunction={representantFunction}
            onSaved={handleSignatureSaved}
          />
          <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
            {["personnel", "representant"].map(s => {
              const sig = getSig(s);
              const expectedName = s === "personnel"
                ? `${selectedPerson.nom || ""} ${selectedPerson.prenom || ""}`.trim()
                : representantName;
              const shownName = sig?.signataireName || expectedName || (s === "personnel" ? "PERSONNEL" : "REPRESENTANT");
              return (
                <div key={s} style={{ flex: 1, padding: "6px 8px", borderRadius: 8, background: sig ? "rgba(111,157,120,0.16)" : "rgba(217,137,106,0.12)", border: `1px solid ${sig ? "rgba(111,157,120,0.35)" : "rgba(217,137,106,0.3)"}`, fontSize: 9, color: sig ? "#2e6a44" : "#8f4a32", textAlign: "center" }}>
                  {shownName}<br />
                  {sig ? `✓ ${new Date(sig.signedAt).toLocaleDateString("fr-FR")}` : "NON SIGNE"}
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}
