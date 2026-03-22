import React, { useState } from "react";

export default function MobilePersonSearch({ persons, selectedPerson, onSelectPerson }) {
  const [search, setSearch] = useState("");
  const [open, setOpen] = useState(false);

  const filtered = persons.filter(p => {
    if (!search) return true;
    const q = search.toLowerCase();
    return (p.nom || "").toLowerCase().includes(q) || (p.prenom || "").toLowerCase().includes(q) || (p.sites || []).join(" ").toLowerCase().includes(q);
  }).slice(0, 10);

  const select = (p) => {
    onSelectPerson(p);
    setSearch("");
    setOpen(false);
  };

  return (
    <div style={{ marginBottom: 8, position: "relative" }}>
      <div style={{ background: "rgba(244,241,234,0.98)", border: "1px solid rgba(173,190,199,0.98)", borderRadius: 11, padding: "8px 10px", boxShadow: "0 4px 12px rgba(31,49,59,0.08)" }}>
        {selectedPerson && (
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
            <div>
              <span style={{ fontWeight: 700, fontSize: 13, color: "#0f1e26" }}>{selectedPerson.nom} {selectedPerson.prenom}</span>
              <span style={{ fontSize: 10, color: "#3f5662", marginLeft: 8 }}>{(selectedPerson.sites || []).join(", ")}</span>
            </div>
            <button onClick={() => onSelectPerson(null)} style={{ fontSize: 10, padding: "2px 8px", borderRadius: 7, border: "1px solid rgba(202,91,96,0.3)", background: "rgba(202,91,96,0.1)", color: "#7d2a31", cursor: "pointer" }}>✕</button>
          </div>
        )}
        <div style={{ position: "relative" }}>
          <input
            type="search"
            value={search}
            onChange={e => { setSearch(e.target.value); setOpen(true); }}
            onFocus={() => setOpen(true)}
            placeholder="RECHERCHER UNE PERSONNE..."
            style={{ width: "100%", padding: "7px 10px 7px 32px", borderRadius: 9, border: "1px solid rgba(173,190,199,0.98)", background: "#fffdfa", fontSize: 12, color: "#0f1e26", boxSizing: "border-box" }}
          />
          <span style={{ position: "absolute", left: 9, top: "50%", transform: "translateY(-50%)", fontSize: 14 }}>🔍</span>
        </div>
      </div>

      {open && search && filtered.length > 0 && (
        <div style={{ position: "absolute", top: "100%", left: 0, right: 0, zIndex: 200, background: "#fff", border: "1px solid rgba(173,190,199,0.98)", borderRadius: 11, boxShadow: "0 8px 20px rgba(31,49,59,0.14)", maxHeight: 220, overflowY: "auto" }}>
          {filtered.map(p => (
            <button key={p.id} onClick={() => select(p)} style={{ display: "block", width: "100%", padding: "10px 14px", border: "none", borderBottom: "1px solid rgba(173,190,199,0.4)", background: "transparent", textAlign: "left", cursor: "pointer", fontSize: 12 }}>
              <span style={{ fontWeight: 700, color: "#0f1e26" }}>{p.nom} {p.prenom}</span>
              <span style={{ fontSize: 10, color: "#3f5662", marginLeft: 8 }}>{(p.sites || []).join(", ")}</span>
            </button>
          ))}
        </div>
      )}
      {open && search && filtered.length === 0 && (
        <div style={{ position: "absolute", top: "100%", left: 0, right: 0, zIndex: 200, background: "#fff", border: "1px solid rgba(173,190,199,0.98)", borderRadius: 11, boxShadow: "0 8px 20px rgba(31,49,59,0.14)", padding: "12px 14px", fontSize: 11, color: "#3f5662" }}>
          AUCUN RESULTAT
        </div>
      )}
    </div>
  );
}