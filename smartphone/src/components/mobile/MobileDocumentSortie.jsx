import React, { useState, useEffect } from "react";
import { db } from "@/lib/db";
import MobilePersonSearch from "./MobilePersonSearch";
import MobileSignatureCanvas from "./MobileSignatureCanvas";
import { getEffectBillingCause as getPersistedBillingCause, getEffectStatus, getReplacementCostValue, normalizeManualStatus } from "@/lib/businessRules";

const card = { background: "rgba(244,241,234,0.98)", border: "1px solid rgba(173,190,199,0.98)", borderRadius: 11, padding: "12px", marginBottom: 8, boxShadow: "0 4px 12px rgba(31,49,59,0.10)" };
const docField = { display: "flex", flexDirection: "column", gap: 3, marginBottom: 8 };
const docLabel = { fontSize: 9, color: "#4a6170", letterSpacing: "0.08em" };
const docValue = { padding: "7px 12px", borderRadius: 9, background: "rgba(251,250,247,0.98)", border: "1px solid rgba(152,177,190,0.9)", fontSize: 12, color: "#0f1e26", minHeight: 32, display: "flex", alignItems: "center" };
function normalizeLabel(value) {
  return String(value || "").replace(/\s+/g, " ").trim().toUpperCase();
}

function normalizeCause(value) {
  const normalized = normalizeLabel(value);
  if (normalized === "CASSE") return "HS";
  if (normalized === "PERDU") return "PERTE";
  if (["DETRUIT", "PERTE", "VOL", "HS", "NON RENDU"].includes(normalized)) return normalized;
  return "";
}

function formatCost(value) {
  const amount = Number(value);
  const safeAmount = Number.isFinite(amount) ? amount : 0;
  return `${safeAmount.toFixed(2).replace(".", ",")}€`;
}

function resolveRepresentativeIdFromSignature(signature, representatives) {
  if (!signature || !Array.isArray(representatives) || representatives.length === 0) {
    return "";
  }
  const directId = String(signature.signataireId || "").trim();
  if (directId && representatives.some((rep) => String(rep.id) === directId)) {
    return directId;
  }
  const signedName = normalizeLabel(signature.signataireName);
  if (!signedName) {
    return "";
  }
  const signedFunction = normalizeLabel(signature.signataireFunction);
  const strictMatch = representatives.find(
    (rep) => normalizeLabel(rep.nom) === signedName && (!signedFunction || normalizeLabel(rep.fonction) === signedFunction)
  );
  if (strictMatch) {
    return String(strictMatch.id);
  }
  const looseMatch = representatives.find((rep) => normalizeLabel(rep.nom) === signedName);
  return looseMatch ? String(looseMatch.id) : "";
}

const STATUT_COLORS = {
  "ACTIF": { bg: "rgba(89,148,117,0.16)", color: "#2f5e43" },
  "RENDU": { bg: "rgba(87,143,106,0.2)", color: "#2c513a" },
  "NON RENDU": { bg: "rgba(224,147,82,0.2)", color: "#8e4d1e" },
  "PERDU": { bg: "rgba(202,91,96,0.19)", color: "#7d2a31" },
  "HS": { bg: "rgba(132,140,149,0.22)", color: "#4a545d" },
  "VOL": { bg: "rgba(181,120,172,0.2)", color: "#6f3d73" },
};

export default function MobileDocumentSortie({ persons, effets, selectedPerson, onSelectPerson, setSaveStatus, onDataChange, representatives = [], pricingRules = [], effetTypes = [] }) {
  const [signatures, setSignatures] = useState([]);
  const [activeSection, setActiveSection] = useState("identite");
  const [representantId, setRepresentantId] = useState("");
  const [localEffets, setLocalEffets] = useState([]);
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState(null);

  useEffect(() => {
    if (selectedPerson) {
      loadSignatures();
      const pe = effets.filter(e => e.personId === selectedPerson.id);
      setLocalEffets(
        pe.map((e) => ({
          ...e,
          statut: normalizeManualStatus(e.statut) || "ACTIF",
          cause: normalizeCause(e.cause || e.causeRemplacement),
          _rendus: false,
        }))
      );
    }
  }, [selectedPerson, effets]);

  useEffect(() => {
    setRepresentantId("");
  }, [selectedPerson?.id]);

  useEffect(() => {
    if (!selectedPerson) return;
    const signedRepresentative = signatures.find((s) => s.signer === "representant");
    const resolvedId = resolveRepresentativeIdFromSignature(signedRepresentative, representatives || []);
    if (!resolvedId) return;
    setRepresentantId((prev) => (prev === resolvedId ? prev : resolvedId));
  }, [selectedPerson?.id, signatures, representatives]);

  const loadSignatures = async () => {
    const sigs = await db.Signature.filter({ personId: selectedPerson.id, docType: "exit" });
    setSignatures(sigs);
  };

  const handleSignatureSaved = async ({ signer } = {}) => {
    await loadSignatures();
    if (signer !== "representant" || !onDataChange || !selectedPerson?.id) return;
    const allSignatures = await db.Signature.filter({ personId: selectedPerson.id, docType: "exit" });
    const personnel = allSignatures.find((entry) => entry.signer === "personnel");
    const representant = allSignatures.find((entry) => entry.signer === "representant");
    if (personnel?.signedAt && representant?.signedAt) {
      await onDataChange();
    }
  };

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
  const todayIso = () => new Date().toISOString().slice(0, 10);

  const toggleRendu = (id) => {
    setLocalEffets(prev =>
      prev.map((e) => {
        if (e.id !== id) return e;
        const nextRendu = !e._rendus;
        return {
          ...e,
          _rendus: nextRendu,
          statut: e.statut,
          dateRetour: nextRendu ? (e.dateRetour || todayIso()) : "",
        };
      })
    );
  };

  const totalEffets = localEffets.length;
  const rendus = localEffets.filter((e) => getEffectStatus(selectedPerson, e) === "RESTITUE" || e._rendus).length;
  const getEffetBillingCause = (effet) => {
    return normalizeCause(getPersistedBillingCause(selectedPerson, effet));
  };
  const getEffetReferenceCost = (effet) => {
    const cause = getEffetBillingCause(effet);
    if (!cause) return 0;
    return getReplacementCostValue(pricingRules, effet?.typeEffet, cause, effet?.designation || "");
  };
  const facturableAmounts = localEffets
    .map((e) => getEffetReferenceCost(e))
    .filter((amount) => amount > 0);
  const totalFacturable = facturableAmounts.reduce((sum, amount) => sum + amount, 0);
  const normalizeCostCause = (cause) => {
    const normalized = normalizeLabel(cause);
    if (normalized === "PERDU") return "PERTE";
    if (normalized === "CASSE") return "HS";
    return normalized;
  };
  const costByType = new Map();
  const causeColumns = Array.from(
    new Set(
      pricingRules
        .map((rule) => normalizeCostCause(rule.cause))
        .filter(Boolean)
    )
  );
  pricingRules.forEach((rule) => {
    const type = normalizeLabel(rule.typeEffet);
    if (!type) return;
    const cause = normalizeCostCause(rule.cause);
    if (!cause) return;
    const amount = Number(rule.montant) || 0;
    const current = costByType.get(type) || {};
    current[cause] = amount;
    costByType.set(type, current);
  });
  const baseTypes = Array.from(new Set(
    [...effetTypes.map((t) => normalizeLabel(t)), ...Array.from(costByType.keys())].filter(Boolean)
  ));
  const clauseRows = baseTypes.map((type) => ({
    id: type,
    label: type,
    byCause: costByType.get(type) || {},
  }));

  const handleSaveEffets = async () => {
    if (!selectedPerson) return;
    setSaving(true);
    setSaveStatus("saving");
    try {
      for (const e of localEffets) {
        await db.Effet.update(e.id, {
          statut: normalizeManualStatus(e.statut) || "ACTIF",
          cause: normalizeCause(e.cause),
          dateRetour: String(e.dateRetour || "").trim() ? String(e.dateRetour) : "",
        });
      }
      setSaveStatus("saved");
      setMsg("MODIFICATIONS SAUVEGARDEES");
      onDataChange();
      setTimeout(() => setMsg(null), 2500);
    } catch (error) {
      console.error("Sortie save effets error:", error);
      setSaveStatus("error");
      setMsg(String(error?.message || "ERREUR DE SAUVEGARDE SUPABASE").toUpperCase());
      setTimeout(() => setMsg(null), 2500);
    } finally {
      setSaving(false);
    }
  };

  return (
    <div style={{ padding: "12px 12px 0" }}>
      <MobilePersonSearch persons={persons} selectedPerson={selectedPerson} onSelectPerson={onSelectPerson} />

      {/* Section tabs */}
      <div style={{ display: "flex", gap: 4, marginBottom: 8, overflowX: "auto" }}>
        {[["identite", "IDENTITE"], ["effets", `EFFETS (${totalEffets})`], ["signatures", "SIGNATURES"]].map(([id, lbl]) => (
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
            <div style={{ fontSize: 9, color: "#556d79", letterSpacing: "0.12em" }}>DOCUMENT DE SORTIE</div>
            <div style={{ fontSize: 14, fontWeight: 700, color: "#14242c" }}>ETAT DE RESTITUTION DES EFFETS</div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 10px" }}>
            {[
              ["NOM", selectedPerson.nom],
              ["PRENOM", selectedPerson.prenom],
              ["FONCTION", selectedPerson.fonction],
              ["TYPE PERSONNEL", selectedPerson.typePersonnel],
              ["DATE DU DOCUMENT", fmt(documentDateIso)],
              ["DATE D'ENTREE", fmt(selectedPerson.dateEntree)],
              ["SORTIE PREVUE", fmt(selectedPerson.dateSortiePrevue)],
              ["SORTIE REELLE", fmt(selectedPerson.dateSortieReelle)],
              ["SITES", (selectedPerson.sites || []).join(", ")],
            ].map(([lbl, val]) => (
              <div key={lbl} style={docField}>
                <span style={docLabel}>{lbl}</span>
                <div style={docValue}>{val || "—"}</div>
              </div>
            ))}
          </div>
          {/* Totaux */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 6, marginTop: 8, padding: "8px", background: "rgba(221,231,235,0.6)", borderRadius: 9 }}>
            <div><div style={{ fontSize: 9, color: "#4a6170" }}>TOTAL</div><div style={{ fontSize: 18, fontWeight: 700 }}>{totalEffets}</div></div>
            <div><div style={{ fontSize: 9, color: "#4a6170" }}>RENDUS</div><div style={{ fontSize: 18, fontWeight: 700, color: "#2f5e43" }}>{rendus}</div></div>
            <div><div style={{ fontSize: 9, color: "#4a6170" }}>FACTURABLES</div><div style={{ fontSize: 18, fontWeight: 700, color: "#8e4d1e" }}>{formatCost(totalFacturable)}</div></div>
          </div>
          {totalFacturable > 0 && (
            <div style={{ marginTop: 8, padding: "8px 12px", borderRadius: 9, background: "linear-gradient(180deg, rgba(232,215,167,0.42) 0%, rgba(244,241,234,0.98) 100%)", border: "1px solid rgba(216,169,104,0.74)", textAlign: "center" }}>
              <div style={{ fontSize: 9, color: "#8a5325" }}>TOTAL FACTURABLE</div>
              <div style={{ fontSize: 22, fontWeight: 800, color: "#8a5325" }}>{formatCost(totalFacturable)}</div>
            </div>
          )}
        </div>
      )}

      {selectedPerson && activeSection === "effets" && (
        <div>
          {msg && <div style={{ padding: "8px 12px", borderRadius: 9, background: "rgba(111,157,120,0.2)", color: "#2e6a44", fontSize: 11, marginBottom: 8 }}>{msg}</div>}
          {localEffets.length === 0 ? (
            <div style={{ ...card, textAlign: "center", color: "#3f5662", fontSize: 11 }}>AUCUN EFFET</div>
          ) : (
            localEffets.map(e => {
              const rawStatus = getEffectStatus(selectedPerson, e);
              const displayStatus = rawStatus === "RESTITUE" ? "RENDU" : rawStatus;
              const sc = STATUT_COLORS[displayStatus] || STATUT_COLORS["ACTIF"];
              const lineCost = getEffetReferenceCost(e);
              return (
                <div key={e.id} style={{ ...card, padding: "10px 12px" }}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 4 }}>
                    <div>
                      <span style={{ fontWeight: 700, fontSize: 12 }}>{e.typeEffet}</span>
                      <span style={{ fontSize: 9, marginLeft: 8, padding: "2px 7px", borderRadius: 99, background: sc.bg, color: sc.color }}>{displayStatus}</span>
                    </div>
                    <span style={{ fontSize: 11, color: "#9b5a2a", fontWeight: 600 }}>{formatCost(lineCost)}</span>
                  </div>
                  <div style={{ fontSize: 11, color: "#213b48", marginBottom: 4 }}>{e.designation || "—"}</div>
                  {/* Statut select */}
                  <select
                    value={e.statut}
                    onChange={(ev) => {
                      const rawChoice = normalizeLabel(ev.target.value);
                      const nextStatut = normalizeManualStatus(ev.target.value) || "ACTIF";
                      const nextCause =
                        nextStatut === "PERDU"
                          ? "PERTE"
                          : nextStatut === "VOL"
                            ? "VOL"
                            : nextStatut === "DETRUIT"
                              ? "DETRUIT"
                              : nextStatut === "HS"
                                ? "HS"
                                : e.cause || "";
                      setLocalEffets((prev) =>
                        prev.map((x) =>
                          x.id === e.id
                            ? {
                                ...x,
                                statut: rawChoice === "RESTITUE" ? x.statut : nextStatut,
                                cause: rawChoice === "RESTITUE" ? x.cause : nextCause,
                                _rendus: rawChoice === "RESTITUE",
                                dateRetour: rawChoice === "RESTITUE" ? (x.dateRetour || todayIso()) : "",
                              }
                            : x
                        )
                      );
                    }}
                    style={{ width: "100%", padding: "6px 8px", borderRadius: 7, border: "1px solid rgba(173,190,199,0.98)", background: "#fffdfa", fontSize: 11, color: "#0f1e26" }}
                  >
                    {[
                      ["ACTIF", "ACTIF"],
                      ["RESTITUE", "RENDU"],
                      ["PERDU", "PERDU"],
                      ["HS", "HS"],
                      ["VOL", "VOL"],
                      ["DETRUIT", "DETRUIT"],
                    ].map(([value, label]) => <option key={value} value={value}>{label}</option>)}
                  </select>
                </div>
              );
            })
          )}
          {localEffets.length > 0 && (
            <button onClick={handleSaveEffets} disabled={saving} style={{ width: "100%", padding: "11px", borderRadius: 9, border: "none", background: "#3f6170", color: "#fff", fontSize: 12, fontWeight: 700, cursor: saving ? "default" : "pointer", marginTop: 4, marginBottom: 8, opacity: saving ? 0.6 : 1 }}>
              {saving ? "SAUVEGARDE..." : "SAUVEGARDER LES STATUTS"}
            </button>
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
          <div style={{ ...card, padding: "7px 10px", marginBottom: 8, background: "linear-gradient(180deg, rgba(232,215,167,0.42) 0%, rgba(244,241,234,0.98) 100%)", border: "1px solid rgba(216,169,104,0.74)" }}>
            <div style={{ fontSize: 9, color: "#8a5325", letterSpacing: "0.08em", marginBottom: 2, fontWeight: 700 }}>TOTAL FACTURABLE</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#8a5325", lineHeight: 1.1 }}>{formatCost(totalFacturable)}</div>
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
                    {causeColumns.map((cause) => (
                      <th key={cause} style={{ textAlign: "right", fontSize: 10, color: "#4a6170", padding: "4px 6px", borderBottom: "1px solid rgba(152,177,190,0.8)" }}>{cause}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {clauseRows.map((row) => (
                    <tr key={row.id}>
                      <td style={{ fontSize: 11, color: "#213b48", padding: "6px", borderBottom: "1px solid rgba(152,177,190,0.5)" }}>{row.label}</td>
                      {causeColumns.map((cause) => (
                        <td key={`${row.id}-${cause}`} style={{ fontSize: 11, color: "#213b48", padding: "6px", textAlign: "right", borderBottom: "1px solid rgba(152,177,190,0.5)" }}>
                          {formatCost(row.byCause?.[cause] || 0)}
                        </td>
                      ))}
                    </tr>
                  ))}
                  {clauseRows.length === 0 && (
                    <tr>
                      <td colSpan={Math.max(2, causeColumns.length + 1)} style={{ fontSize: 11, color: "#3f5662", padding: "8px 6px", textAlign: "center" }}>AUCUN EFFET</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
          <MobileSignatureCanvas
            personId={selectedPerson.id}
            docType="exit"
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
            docType="exit"
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
