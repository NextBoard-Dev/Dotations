import React, { useState, useEffect } from "react";
import { db } from "@/lib/db";
import { normalizeManualStatus } from "@/lib/businessRules";

const inputStyle = { padding: "8px 10px", borderRadius: 9, border: "1px solid rgba(173,190,199,0.98)", background: "#fffdfa", fontSize: 12, color: "#0f1e26", width: "100%", boxSizing: "border-box" };
const labelStyle = { fontSize: 9, color: "#4a6170", letterSpacing: "0.08em", display: "block", marginBottom: 3 };
const fieldStyle = { marginBottom: 10 };

function normalizeCause(value) {
  const normalized = String(value || "").trim().toUpperCase();
  if (normalized === "CASSE") return "HS";
  if (normalized === "PERDU") return "PERTE";
  if (["DETRUIT", "PERTE", "VOL", "HS"].includes(normalized)) return normalized;
  return "";
}

function inferCauseFromStatus(value) {
  const normalizedStatus = normalizeManualStatus(value);
  if (normalizedStatus === "PERDU") return "PERTE";
  if (normalizedStatus === "DETRUIT") return "DETRUIT";
  if (normalizedStatus === "VOL") return "VOL";
  if (normalizedStatus === "HS") return "HS";
  return "";
}

function MobileEffetForm({ personId, editingEffet, onSaved, onCancel, setSaveStatus, bases = {} }) {
  const [form, setForm] = useState({ typeEffet: "", designation: "", siteReference: "", numeroIdentification: "", vehiculeImmatriculation: "", dateRemise: "", dateRetour: "", statut: "ACTIF", cause: "", dateRemplacement: "", coutRemplacement: "", commentaire: "" });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const typesEffetsRef = Array.from(
    new Set([
      ...(Array.isArray(bases.typesEffets) ? bases.typesEffets.map((entry) => String(entry || "").trim()).filter(Boolean) : []),
      String(form.typeEffet || "").trim(),
    ].filter(Boolean))
  );
  const statutsRef = Array.from(
    new Set([
      ...(Array.isArray(bases.statutsObjetManuels) ? bases.statutsObjetManuels.map((s) => normalizeManualStatus(s)).filter(Boolean) : []),
      normalizeManualStatus(form.statut) || "ACTIF",
    ].filter(Boolean))
  );
  const sitesRef = Array.isArray(bases.sites) ? bases.sites : [];
  const referencesEffetsRef = Array.isArray(bases.referencesEffets) ? bases.referencesEffets : [];
  const normalizedType = String(form.typeEffet || "").trim().toUpperCase();
  const normalizedSite = String(form.siteReference || "").trim().toUpperCase();
  const designationOptions = Array.from(
    new Set(
      referencesEffetsRef
        .filter((entry) => {
          const refType = String(entry?.typeEffet || "").trim().toUpperCase();
          const refSite = String(entry?.site || "").trim().toUpperCase();
          if (normalizedType && refType && refType !== normalizedType) return false;
          if (normalizedType === "CLE" && normalizedSite && refSite && refSite !== normalizedSite) return false;
          return true;
        })
        .map((entry) => String(entry?.designation || "").trim())
        .filter(Boolean)
    )
  );
  const siteOptions = Array.from(new Set((sitesRef || []).map((entry) => String(entry || "").trim()).filter(Boolean)));
  const siteValueInList = !form.siteReference || siteOptions.includes(form.siteReference);
  const designationValueInList = !form.designation || designationOptions.includes(form.designation);

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
        statut: normalizeManualStatus(editingEffet.statut) || "ACTIF",
        cause: normalizeCause(editingEffet.cause || editingEffet.causeRemplacement),
        dateRemplacement: editingEffet.dateRemplacement || "",
        coutRemplacement: editingEffet.coutRemplacement || "",
        commentaire: editingEffet.commentaire || "",
      });
    } else {
      setForm({ typeEffet: "", designation: "", siteReference: "", numeroIdentification: "", vehiculeImmatriculation: "", dateRemise: new Date().toISOString().slice(0, 10), dateRetour: "", statut: "ACTIF", cause: "", dateRemplacement: "", coutRemplacement: "", commentaire: "" });
    }
  }, [editingEffet]);

  useEffect(() => {
    if (!form.designation) return;
    if (!designationOptions.includes(form.designation)) {
      setForm((prev) => ({ ...prev, designation: "" }));
    }
  }, [form.typeEffet, form.siteReference]);

  const handleSave = async () => {
    if (!form.typeEffet) { setError("TYPE D'EFFET OBLIGATOIRE"); return; }
    setSaving(true);
    setSaveStatus("saving");
    setError(null);
    try {
      const data = {
        ...form,
        statut: normalizeManualStatus(form.statut) || "ACTIF",
        cause: normalizeCause(form.cause) || inferCauseFromStatus(form.statut),
        personId,
        coutRemplacement: form.coutRemplacement ? parseFloat(String(form.coutRemplacement).replace(",", ".")) : null
      };
      if (editingEffet) {
        await db.Effet.update(editingEffet.id, data);
      } else {
        await db.Effet.create(data);
      }
      setSaveStatus("saved");
      onSaved();
    } catch (error) {
      console.error("Effet save error:", error);
      setSaveStatus("error");
      setError(String(error?.message || "ERREUR DE SAUVEGARDE SUPABASE").toUpperCase());
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
    } catch (error) {
      console.error("Effet delete error:", error);
      setSaveStatus("error");
      setError(String(error?.message || "ERREUR DE SAUVEGARDE SUPABASE").toUpperCase());
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
          <label style={labelStyle}>SITE REFERENCE</label>
          <select value={form.siteReference} onChange={e => setForm(f => ({ ...f, siteReference: e.target.value }))} style={inputStyle}>
            <option value="">{siteOptions.length ? "SELECTIONNER UN SITE" : "AUCUN SITE DISPONIBLE"}</option>
            {!siteValueInList && <option value={form.siteReference}>{form.siteReference}</option>}
            {siteOptions.map((site) => <option key={site} value={site}>{site}</option>)}
          </select>
        </div>
      </div>

      <div style={fieldStyle}>
        <label style={labelStyle}>DESIGNATION</label>
        <select value={form.designation} onChange={e => setForm(f => ({ ...f, designation: e.target.value }))} style={inputStyle}>
          <option value="">
            {designationOptions.length ? "SELECTIONNER UNE DESIGNATION" : "AUCUNE DESIGNATION DISPONIBLE"}
          </option>
          {!designationValueInList && <option value={form.designation}>{form.designation}</option>}
          {designationOptions.map((d) => <option key={d} value={d}>{d}</option>)}
        </select>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 10px" }}>
        <div style={fieldStyle}>
          <label style={labelStyle}>N° IDENTIFICATION</label>
          <input value={form.numeroIdentification} onChange={e => setForm(f => ({ ...f, numeroIdentification: e.target.value }))} style={inputStyle} />
        </div>
        <div style={fieldStyle}>
          <label style={labelStyle}>STATUT</label>
          <select value={form.statut} onChange={e => setForm(f => ({ ...f, statut: e.target.value }))} style={inputStyle}>
            {statutsRef.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
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
