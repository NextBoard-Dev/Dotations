import React, { useState, useEffect } from "react";
import { db } from "@/lib/db";

const inputStyle = { padding: "8px 10px", borderRadius: 9, border: "1px solid rgba(173,190,199,0.98)", background: "#fffdfa", fontSize: 12, color: "#0f1e26", width: "100%", boxSizing: "border-box" };
const labelStyle = { fontSize: 9, color: "#4a6170", letterSpacing: "0.08em", display: "block", marginBottom: 3 };
const fieldStyle = { marginBottom: 10 };

const TYPES_EFFET = ["CLE", "BADGE", "CARTE", "TELECOMMANDE", "AUTRE"];
const STATUTS = ["ACTIF", "RESTITUE", "NON RENDU", "PERDU", "HS", "VOLE", "DETRUIT"];

function MobileEffetForm({ personId, editingEffet, onSaved, onCancel, setSaveStatus, bases = {} }) {
  const [form, setForm] = useState({ typeEffet: "", designation: "", siteReference: "", numeroIdentification: "", vehiculeImmatriculation: "", dateRemise: "", dateRetour: "", statut: "ACTIF", dateRemplacement: "", coutRemplacement: "", commentaire: "" });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const typesEffetsRef = Array.isArray(bases.typesEffets) && bases.typesEffets.length ? bases.typesEffets : TYPES_EFFET;
  const statutsRef = Array.isArray(bases.statutsObjetManuels) && bases.statutsObjetManuels.length
    ? Array.from(new Set(["ACTIF", "RESTITUE", "NON RENDU", "PERDU", "HS", "VOLE", "DETRUIT", ...bases.statutsObjetManuels]))
    : STATUTS;
  const sitesRef = Array.isArray(bases.sites) ? bases.sites : [];

  useEffect(() => {
    if (editingEffet) {
      setForm({
        typeEffet: editingEffet.typeEffet || "",
        designation: editingEffet.designation || "",
        siteReference: editingEffet.siteReference || "",
        numeroIdentification: editingEffet.numeroIdentification || "",
        vehiculeImmatriculation: editingEffet.vehiculeImmatriculation || "",
        dateRemise: editingEffet.dateRemise || "",
        dateRetour: editingEffet.dateRetour || "",
        statut: editingEffet.statut || "ACTIF",
        dateRemplacement: editingEffet.dateRemplacement || "",
        coutRemplacement: editingEffet.coutRemplacement || "",
        commentaire: editingEffet.commentaire || "",
      });
    } else {
      setForm({ typeEffet: "", designation: "", siteReference: "", numeroIdentification: "", vehiculeImmatriculation: "", dateRemise: new Date().toISOString().slice(0, 10), dateRetour: "", statut: "ACTIF", dateRemplacement: "", coutRemplacement: "", commentaire: "" });
    }
  }, [editingEffet]);

  const handleSave = async () => {
    if (!form.typeEffet) { setError("TYPE D'EFFET OBLIGATOIRE"); return; }
    setSaving(true);
    setSaveStatus("saving");
    setError(null);
    try {
      const data = { ...form, personId, coutRemplacement: form.coutRemplacement ? parseFloat(String(form.coutRemplacement).replace(",", ".")) : null };
      if (editingEffet) {
        await db.Effet.update(editingEffet.id, data);
      } else {
        await db.Effet.create(data);
      }
      setSaveStatus("saved");
      onSaved();
    } catch {
      setSaveStatus("saved");
      setError("SAUVEGARDE SUPABASE TEMPORAIREMENT BLOQUEE");
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async () => {
    if (!editingEffet) return;
    if (!window.confirm("Supprimer cet effet ?")) return;
    try {
      setSaveStatus("saving");
      await db.Effet.delete(editingEffet.id);
      setSaveStatus("saved");
      onSaved();
    } catch {
      setSaveStatus("saved");
      setError("SAUVEGARDE SUPABASE TEMPORAIREMENT BLOQUEE");
    }
  };

  return (
    <div style={{ background: "rgba(244,241,234,0.98)", border: "1px solid rgba(173,190,199,0.98)", borderRadius: 11, padding: "12px", boxShadow: "0 4px 12px rgba(31,49,59,0.10)" }}>
      {error && <div style={{ padding: "6px 10px", borderRadius: 7, background: "rgba(202,91,96,0.12)", color: "#7d2a31", fontSize: 11, marginBottom: 8 }}>{error}</div>}

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 10px" }}>
        <div style={fieldStyle}>
          <label style={labelStyle}>TYPE D'EFFET *</label>
          <select value={form.typeEffet} onChange={e => setForm(f => ({ ...f, typeEffet: e.target.value }))} style={inputStyle}>
            <option value="">SELECTIONNER</option>
            {typesEffetsRef.map(t => <option key={t} value={t}>{t}</option>)}
          </select>
        </div>
        <div style={fieldStyle}>
          <label style={labelStyle}>STATUT</label>
          <select value={form.statut} onChange={e => setForm(f => ({ ...f, statut: e.target.value }))} style={inputStyle}>
            {statutsRef.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </div>
      </div>

      <div style={fieldStyle}>
        <label style={labelStyle}>DESIGNATION</label>
        <input value={form.designation} onChange={e => setForm(f => ({ ...f, designation: e.target.value }))} style={inputStyle} placeholder="DESIGNATION / REFERENCE" />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 10px" }}>
        <div style={fieldStyle}>
          <label style={labelStyle}>N° IDENTIFICATION</label>
          <input value={form.numeroIdentification} onChange={e => setForm(f => ({ ...f, numeroIdentification: e.target.value }))} style={inputStyle} />
        </div>
        <div style={fieldStyle}>
          <label style={labelStyle}>SITE REFERENCE</label>
          <input value={form.siteReference} onChange={e => setForm(f => ({ ...f, siteReference: e.target.value }))} style={inputStyle} />
        </div>
        <div style={fieldStyle}>
          <label style={labelStyle}>DATE REMISE</label>
          <input type="date" value={form.dateRemise} onChange={e => setForm(f => ({ ...f, dateRemise: e.target.value }))} style={inputStyle} />
        </div>
        <div style={fieldStyle}>
          <label style={labelStyle}>DATE RETOUR</label>
          <input type="date" value={form.dateRetour} onChange={e => setForm(f => ({ ...f, dateRetour: e.target.value }))} style={inputStyle} />
        </div>
        <div style={fieldStyle}>
          <label style={labelStyle}>DATE REMPLACEMENT</label>
          <input type="date" value={form.dateRemplacement} onChange={e => setForm(f => ({ ...f, dateRemplacement: e.target.value }))} style={inputStyle} />
        </div>
        <div style={fieldStyle}>
          <label style={labelStyle}>COUT REMPLACEMENT (€)</label>
          <input value={form.coutRemplacement} onChange={e => setForm(f => ({ ...f, coutRemplacement: e.target.value }))} style={inputStyle} inputMode="decimal" placeholder="0,00" />
        </div>
      </div>

      <div style={fieldStyle}>
        <label style={labelStyle}>COMMENTAIRE</label>
        <textarea value={form.commentaire} onChange={e => setForm(f => ({ ...f, commentaire: e.target.value }))} style={{ ...inputStyle, minHeight: 60, resize: "vertical" }} rows={3} />
      </div>

      <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
        <button onClick={handleSave} disabled={saving} style={{ flex: 1, minWidth: 100, padding: "9px 10px", borderRadius: 9, border: "none", background: "#3f6170", color: "#fff", fontSize: 11, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.6 : 1 }}>
          {editingEffet ? "MODIFIER" : "AJOUTER"}
        </button>
        <button onClick={onCancel} style={{ flex: 1, minWidth: 80, padding: "9px 10px", borderRadius: 9, border: "1px solid rgba(63,97,112,0.3)", background: "rgba(63,97,112,0.1)", color: "#213b48", fontSize: 11, cursor: "pointer" }}>
          ANNULER
        </button>
        {editingEffet && (
          <button onClick={handleDelete} style={{ padding: "9px 12px", borderRadius: 9, border: "1px solid rgba(202,91,96,0.4)", background: "rgba(202,91,96,0.1)", color: "#7d2a31", fontSize: 11, cursor: "pointer" }}>
            SUPPRIMER
          </button>
        )}
      </div>
    </div>
  );
}

export { MobileEffetForm };
export default MobileEffetForm;
