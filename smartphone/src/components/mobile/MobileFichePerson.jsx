import React, { useState, useEffect } from "react";
import { db } from "@/lib/db";
import MobilePersonSearch from "./MobilePersonSearch";
import MobileEffetForm from "./MobileEffetForm";
import MobileEffetsList from "./MobileEffetsList";
import { getDossierStatus } from "@/lib/businessRules";

const card = { background: "rgba(244,241,234,0.98)", border: "1px solid rgba(173,190,199,0.98)", borderRadius: 11, padding: "12px", marginBottom: 8, boxShadow: "0 4px 12px rgba(31,49,59,0.10)" };
const fieldStyle = { display: "flex", flexDirection: "column", gap: 3, marginBottom: 8 };
const labelStyle = { fontSize: 9, color: "#4a6170", letterSpacing: "0.08em" };
const inputStyle = { padding: "7px 9px", borderRadius: 9, border: "1px solid rgba(173,190,199,0.98)", background: "#fffdfa", fontSize: 12, color: "#0f1e26", width: "100%", boxSizing: "border-box" };

const STATUTS = ["EN POSTE", "SORTIE PREVUE", "SORTI"];
export default function MobileFichePerson({ persons, effets, selectedPerson, onSelectPerson, onDataChange, setSaveStatus, onNavigate, bases = {} }) {
  const [form, setForm] = useState({ nom: "", prenom: "", fonction: "", sites: [], typePersonnel: "", typeContrat: "", dateEntree: "", dateSortiePrevue: "", dateSortieReelle: "" });
  const [lastSavedForm, setLastSavedForm] = useState(null);
  const [editingEffet, setEditingEffet] = useState(null);
  const [activeSection, setActiveSection] = useState("infos"); // "infos" | "effets" | "add-effet"
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState(null);
  const fonctionsRef = Array.from(
    new Set([
      ...(Array.isArray(bases.fonctions) ? bases.fonctions.map((entry) => String(entry || "").trim()).filter(Boolean) : []),
      String(form.fonction || "").trim(),
    ].filter(Boolean))
  );
  const typesPersonnelRef = Array.from(
    new Set([
      ...(Array.isArray(bases.typesPersonnel) ? bases.typesPersonnel.map((entry) => String(entry || "").trim()).filter(Boolean) : []),
      String(form.typePersonnel || "").trim(),
    ].filter(Boolean))
  );
  const typesContratsRef = Array.from(
    new Set([
      ...(Array.isArray(bases.typesContrats) ? bases.typesContrats.map((entry) => String(entry || "").trim()).filter(Boolean) : []),
      String(form.typeContrat || "").trim(),
    ].filter(Boolean))
  );
  const sitesRef = Array.isArray(bases.sites) ? bases.sites : [];

  useEffect(() => {
    if (selectedPerson) {
      const nextForm = {
        nom: selectedPerson.nom || "",
        prenom: selectedPerson.prenom || "",
        fonction: selectedPerson.fonction || "",
        sites: selectedPerson.sites || [],
        typePersonnel: selectedPerson.typePersonnel || "",
        typeContrat: selectedPerson.typeContrat || "",
        dateEntree: selectedPerson.dateEntree || "",
        dateSortiePrevue: selectedPerson.dateSortiePrevue || "",
        dateSortieReelle: selectedPerson.dateSortieReelle || "",
      };
      setForm(nextForm);
      setLastSavedForm(nextForm);
    } else {
      setLastSavedForm(null);
    }
  }, [selectedPerson]);

  const personEffets = selectedPerson ? effets.filter(e => e.personId === selectedPerson.id) : [];

  const handleSavePerson = async () => {
    setSaving(true);
    setSaveStatus("saving");
    try {
      if (selectedPerson) {
        await db.Person.update(selectedPerson.id, form);
        setMsg("MODIFICATIONS ENREGISTREES");
      } else {
        const created = await db.Person.create(form);
        onSelectPerson(created);
        setMsg("PERSONNE AJOUTEE");
      }
      setLastSavedForm(form);
      setSaveStatus("saved");
      onDataChange();
    } catch (error) {
      console.error("Person save error:", error);
      setMsg(String(error?.message || "ERREUR DE SAUVEGARDE SUPABASE").toUpperCase());
      setSaveStatus("error");
    } finally {
      setSaving(false);
      setTimeout(() => setMsg(null), 2500);
    }
  };

  const handleDeletePerson = async () => {
    if (!selectedPerson) return;
    if (!window.confirm(`Supprimer ${selectedPerson.nom} ${selectedPerson.prenom} ?`)) return;
    try {
      setSaveStatus("saving");
      for (const e of personEffets) await db.Effet.delete(e.id);
      await db.Person.delete(selectedPerson.id);
      onSelectPerson(null);
      setSaveStatus("saved");
      onDataChange();
      setForm({ nom: "", prenom: "", fonction: "", sites: [], typePersonnel: "", typeContrat: "", dateEntree: "", dateSortiePrevue: "", dateSortieReelle: "" });
    } catch (error) {
      console.error("Person delete error:", error);
      setSaveStatus("error");
      setMsg(String(error?.message || "ERREUR DE SAUVEGARDE SUPABASE").toUpperCase());
      setTimeout(() => setMsg(null), 2500);
    }
  };

  const computeStatut = () => {
    return getDossierStatus(form);
  };

  const normalizeSites = (value) => (Array.isArray(value) ? value.map((entry) => String(entry || "").trim()) : []);
  const hasUnsavedChanges = (() => {
    if (!lastSavedForm) {
      const hasData = Object.values(form || {}).some((entry) => {
        if (Array.isArray(entry)) return entry.length > 0;
        return String(entry || "").trim() !== "";
      });
      return hasData;
    }
    return (
      String(form.nom || "") !== String(lastSavedForm.nom || "") ||
      String(form.prenom || "") !== String(lastSavedForm.prenom || "") ||
      String(form.fonction || "") !== String(lastSavedForm.fonction || "") ||
      String(form.typePersonnel || "") !== String(lastSavedForm.typePersonnel || "") ||
      String(form.typeContrat || "") !== String(lastSavedForm.typeContrat || "") ||
      String(form.dateEntree || "") !== String(lastSavedForm.dateEntree || "") ||
      String(form.dateSortiePrevue || "") !== String(lastSavedForm.dateSortiePrevue || "") ||
      String(form.dateSortieReelle || "") !== String(lastSavedForm.dateSortieReelle || "") ||
      JSON.stringify(normalizeSites(form.sites)) !== JSON.stringify(normalizeSites(lastSavedForm.sites))
    );
  })();

  useEffect(() => {
    if (saving) return;
    setSaveStatus(hasUnsavedChanges ? "unsaved" : "saved");
  }, [hasUnsavedChanges, saving, setSaveStatus]);

  const toggleSite = (site) => {
    setForm((prev) => {
      const current = Array.isArray(prev.sites) ? prev.sites : [];
      const exists = current.includes(site);
      return {
        ...prev,
        sites: exists ? current.filter((s) => s !== site) : [...current, site],
      };
    });
  };

  return (
    <div style={{ padding: "12px 12px 0" }}>
      <MobilePersonSearch persons={persons} selectedPerson={selectedPerson} onSelectPerson={p => { onSelectPerson(p); setActiveSection("infos"); }} />

      {/* Section tabs */}
      <div style={{ display: "flex", gap: 4, marginBottom: 8 }}>
        {[["infos", "INFOS"], ["effets", `EFFETS (${personEffets.length})`], ["add-effet", editingEffet ? "MODIFIER EFFET" : "AJOUTER EFFET"]].map(([id, lbl]) => (
          <button key={id} onClick={() => setActiveSection(id)} style={{ flex: 1, padding: "7px 4px", borderRadius: 8, border: "1px solid rgba(63,97,112,0.25)", background: activeSection === id ? "rgba(63,97,112,0.22)" : "rgba(63,97,112,0.08)", color: activeSection === id ? "#213b48" : "#3f5662", fontSize: 9, letterSpacing: "0.04em", fontWeight: activeSection === id ? 700 : 400, cursor: "pointer" }}>
            {lbl}
          </button>
        ))}
      </div>

      {msg && <div style={{ padding: "8px 12px", borderRadius: 9, background: "rgba(111,157,120,0.2)", color: "#2e6a44", fontSize: 11, marginBottom: 8, border: "1px solid rgba(111,157,120,0.35)" }}>{msg}</div>}

      {activeSection === "infos" && (
        <div style={card}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 10px" }}>
            {[
              ["NOM", "nom", "text"],
              ["PRENOM", "prenom", "text"],
              ["DATE D'ENTREE", "dateEntree", "date"],
              ["DATE SORTIE PREVUE", "dateSortiePrevue", "date"],
              ["DATE SORTIE REELLE", "dateSortieReelle", "date"],
            ].map(([lbl, key, type]) => (
              <div key={key} style={fieldStyle}>
                <span style={labelStyle}>{lbl}</span>
                <input type={type} value={form[key] || ""} onChange={e => setForm(f => ({ ...f, [key]: e.target.value }))} style={inputStyle} />
              </div>
            ))}
            <div style={fieldStyle}>
              <span style={labelStyle}>FONCTION</span>
              <select value={form.fonction} onChange={e => setForm(f => ({ ...f, fonction: e.target.value }))} style={inputStyle}>
                <option value="">SELECTIONNER</option>
                {fonctionsRef.map(f => <option key={f} value={f}>{f}</option>)}
              </select>
            </div>
            <div style={fieldStyle}>
              <span style={labelStyle}>TYPE PERSONNEL</span>
              <select value={form.typePersonnel} onChange={e => setForm(f => ({ ...f, typePersonnel: e.target.value }))} style={inputStyle}>
                <option value="">SELECTIONNER</option>
                {typesPersonnelRef.map(t => <option key={t} value={t}>{t}</option>)}
              </select>
            </div>
            <div style={fieldStyle}>
              <span style={labelStyle}>TYPE CONTRAT</span>
              <select value={form.typeContrat} onChange={e => setForm(f => ({ ...f, typeContrat: e.target.value }))} style={inputStyle}>
                <option value="">SELECTIONNER</option>
                {typesContratsRef.map(t => <option key={t} value={t}>{t}</option>)}
              </select>
            </div>
          </div>
          <div style={fieldStyle}>
            <span style={labelStyle}>SITES (SELECTION MULTIPLE)</span>
            <div style={{ border: "1px solid rgba(173,190,199,0.98)", borderRadius: 9, background: "#fffdfa", padding: "8px" }}>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
                {sitesRef.map((site) => {
                  const checked = (form.sites || []).includes(site);
                  return (
                    <button
                      key={site}
                      type="button"
                      onClick={() => toggleSite(site)}
                      style={{
                        minHeight: 32,
                        padding: "6px 8px",
                        borderRadius: 8,
                        border: "1px solid rgba(63,97,112,0.25)",
                        background: checked ? "rgba(63,97,112,0.22)" : "rgba(63,97,112,0.08)",
                        color: checked ? "#213b48" : "#3f5662",
                        fontSize: 10,
                        textAlign: "left",
                        cursor: "pointer",
                      }}
                    >
                      {checked ? "✓ " : ""}{site}
                    </button>
                  );
                })}
                {sitesRef.length === 0 && (
                  <span style={{ fontSize: 10, color: "#556d79" }}>AUCUN SITE DISPONIBLE</span>
                )}
              </div>
            </div>
          </div>
          <div style={fieldStyle}>
            <span style={labelStyle}>STATUT DOSSIER (AUTO)</span>
            <input readOnly value={computeStatut()} style={{ ...inputStyle, background: "rgba(63,97,112,0.12)", fontWeight: 600 }} />
          </div>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 4 }}>
            <button onClick={handleSavePerson} disabled={saving} style={{ flex: 1, minWidth: 120, padding: "9px 10px", borderRadius: 9, border: "none", background: "#3f6170", color: "#fff", fontSize: 11, fontWeight: 700, letterSpacing: "0.06em", cursor: "pointer" }}>
              {selectedPerson ? "ENREGISTRER" : "AJOUTER"}
            </button>
            {selectedPerson && (
              <>
                <button onClick={() => onNavigate("arrivee", selectedPerson)} style={{ flex: 1, minWidth: 100, padding: "9px 8px", borderRadius: 9, border: "1px solid rgba(63,97,112,0.3)", background: "rgba(63,97,112,0.12)", color: "#213b48", fontSize: 10, cursor: "pointer" }}>
                  ENTREE / SIT.
                </button>
                <button onClick={() => onNavigate("sortie", selectedPerson)} style={{ flex: 1, minWidth: 100, padding: "9px 8px", borderRadius: 9, border: "1px solid rgba(63,97,112,0.3)", background: "rgba(63,97,112,0.12)", color: "#213b48", fontSize: 10, cursor: "pointer" }}>
                  SORTIE / SIT.
                </button>
                <button onClick={handleDeletePerson} style={{ padding: "9px 12px", borderRadius: 9, border: "1px solid rgba(202,91,96,0.4)", background: "rgba(202,91,96,0.12)", color: "#7d2a31", fontSize: 10, cursor: "pointer" }}>
                  SUPPRIMER
                </button>
              </>
            )}
          </div>
        </div>
      )}

      {activeSection === "effets" && (
        <MobileEffetsList
          effets={personEffets}
          onEdit={effet => { setEditingEffet(effet); setActiveSection("add-effet"); }}
        />
      )}

      {activeSection === "add-effet" && selectedPerson && (
        <MobileEffetForm
          bases={bases}
          personId={selectedPerson.id}
          editingEffet={editingEffet}
          onSaved={() => { setEditingEffet(null); setActiveSection("effets"); onDataChange(); setSaveStatus("saved"); }}
          onCancel={() => { setEditingEffet(null); setActiveSection("effets"); }}
          setSaveStatus={setSaveStatus}
        />
      )}
      {activeSection === "add-effet" && !selectedPerson && (
        <div style={{ ...card, textAlign: "center", color: "#3f5662", fontSize: 11 }}>SELECTIONNEZ UNE PERSONNE D'ABORD</div>
      )}
    </div>
  );
}
