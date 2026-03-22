import React, { useState } from "react";

const card = { background: "rgba(244,241,234,0.98)", border: "1px solid rgba(173,190,199,0.98)", borderRadius: 11, padding: "10px 12px", marginBottom: 8, boxShadow: "0 4px 12px rgba(31,49,59,0.10)" };
const label = { fontSize: 9, color: "#4a6170", letterSpacing: "0.08em", margin: "0 0 2px" };
const value = { fontSize: 22, fontWeight: 700, color: "#0f1e26", margin: 0, lineHeight: 1 };

function statusColor(s) {
  if (s === "EN POSTE") return { bg: "rgba(89,148,117,0.16)", color: "#2f5e43", border: "rgba(89,148,117,0.38)" };
  if (s === "SORTIE PREVUE") return { bg: "rgba(224,147,82,0.2)", color: "#8e4d1e", border: "rgba(224,147,82,0.42)" };
  if (s === "SORTI") return { bg: "rgba(202,91,96,0.19)", color: "#7d2a31", border: "rgba(202,91,96,0.42)" };
  return { bg: "rgba(93,120,134,0.12)", color: "#213b48", border: "rgba(93,120,134,0.3)" };
}

export default function MobileOverview({ persons, effets, onSelectPerson }) {
  const [search, setSearch] = useState("");

  const filtered = persons.filter(p => {
    if (!search) return true;
    const q = search.toLowerCase();
    return (p.nom || "").toLowerCase().includes(q) ||
      (p.prenom || "").toLowerCase().includes(q) ||
      (p.sites || []).join(" ").toLowerCase().includes(q);
  });

  const totalEffets = effets.length;
  const nonRendus = effets.filter(e => ["NON RENDU", "PERDU", "VOLE"].includes(e.statut)).length;
  const enPoste = persons.filter(p => p.statutDossier !== "SORTI").length;

  return (
    <div style={{ padding: "12px 12px 0" }}>
      {/* KPIs */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 6, marginBottom: 10 }}>
        {[
          { label: "EN POSTE", value: enPoste },
          { label: "EFFETS CONFIES", value: totalEffets },
          { label: "NON RENDUS", value: nonRendus },
        ].map(k => (
          <div key={k.label} style={{ ...card, padding: "8px 10px", marginBottom: 0, display: "flex", flexDirection: "column" }}>
            <p style={label}>{k.label}</p>
            <p style={value}>{k.value}</p>
          </div>
        ))}
      </div>

      {/* Search */}
      <div style={{ ...card, padding: "8px 10px", marginBottom: 8 }}>
        <input
          type="search"
          value={search}
          onChange={e => setSearch(e.target.value)}
          placeholder="RECHERCHER UN NOM, PRENOM, SITE..."
          style={{ width: "100%", padding: "7px 10px", borderRadius: 9, border: "1px solid rgba(173,190,199,0.98)", background: "#fffdfa", fontSize: 12, color: "#0f1e26", boxSizing: "border-box" }}
        />
      </div>

      {/* Person list */}
      <div>
        {filtered.length === 0 && (
          <div style={{ textAlign: "center", color: "#3f5662", fontSize: 11, padding: 20 }}>AUCUN RESULTAT</div>
        )}
        {filtered.map(p => {
          const personEffets = effets.filter(e => e.personId === p.id);
          const sc = statusColor(p.statutDossier);
          return (
            <button
              key={p.id}
              onClick={() => onSelectPerson(p)}
              style={{ ...card, width: "100%", textAlign: "left", cursor: "pointer", display: "flex", alignItems: "center", gap: 10, padding: "10px 12px" }}
            >
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontWeight: 700, fontSize: 13, color: "#0f1e26", marginBottom: 2 }}>{p.nom} {p.prenom}</div>
                <div style={{ fontSize: 10, color: "#3f5662" }}>{(p.sites || []).join(", ")} • {p.typePersonnel || "—"}</div>
                <div style={{ fontSize: 10, color: "#3f5662", marginTop: 2 }}>{personEffets.length} effet(s)</div>
              </div>
              <span style={{ fontSize: 9, padding: "3px 8px", borderRadius: 99, background: sc.bg, color: sc.color, border: `1px solid ${sc.border}`, whiteSpace: "nowrap" }}>
                {p.statutDossier || "—"}
              </span>
            </button>
          );
        })}
      </div>
    </div>
  );
}