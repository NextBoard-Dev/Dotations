import React, { useState } from "react";
import { getDossierStatus, getEffectStatus } from "@/lib/businessRules";

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
  const todayIso = new Date().toISOString().slice(0, 10);

  const filtered = persons.filter(p => {
    if (!search) return true;
    const q = search.toLowerCase();
    return (p.nom || "").toLowerCase().includes(q) ||
      (p.prenom || "").toLowerCase().includes(q) ||
      (p.sites || []).join(" ").toLowerCase().includes(q);
  });

  const personById = new Map((persons || []).map((p) => [String(p.id), p]));
  const totalEffets = effets.length;
  const nonRendus = effets.filter((e) => {
    const person = personById.get(String(e.personId));
    return getEffectStatus(person, e) === "NON RENDU";
  }).length;
  const enPoste = persons.filter((p) => getDossierStatus(p) !== "SORTI").length;
  const alerts = persons.flatMap((p) => {
    const personAlerts = [];
    const sortiePrevue = String(p.dateSortiePrevue || "");
    const sortieReelle = String(p.dateSortieReelle || "");
    if (sortiePrevue && !sortieReelle && sortiePrevue < todayIso) {
      personAlerts.push({ key: `${p.id}-late`, personId: p.id, type: "dateSortiePrevue", text: `${p.nom} ${p.prenom} : SORTIE PREVUE DEPASSEE` });
    }
    if (sortieReelle && sortieReelle <= todayIso) {
      personAlerts.push({ key: `${p.id}-out`, personId: p.id, type: "dateSortieReelle", text: `${p.nom} ${p.prenom} : PERSONNE SORTIE` });
    }
    const nonRendusCount = effets.filter(
      (e) => String(e.personId) === String(p.id) && getEffectStatus(p, e) === "NON RENDU"
    ).length;
    if (nonRendusCount > 0) {
      personAlerts.push({ key: `${p.id}-nr`, personId: p.id, type: "dateSortiePrevue", text: `${p.nom} ${p.prenom} : ${nonRendusCount} EFFET(S) NON RENDU(S)` });
    }
    return personAlerts;
  });
  const visibleAlerts = alerts.slice(0, 3);

  const alertStyle = (type) => {
    if (type === "dateSortieReelle") {
      return {
        bg: "linear-gradient(90deg, rgba(239, 147, 147, 0.18) 0%, rgba(226, 111, 111, 0.08) 100%)",
        color: "#8f2d2d",
        border: "rgba(208, 86, 86, 0.3)",
        borderLeft: "rgba(198, 45, 45, 0.95)",
      };
    }
    return {
      bg: "linear-gradient(90deg, rgba(248, 223, 160, 0.22) 0%, rgba(246, 205, 120, 0.08) 100%)",
      color: "#8b5a1d",
      border: "rgba(223, 173, 67, 0.3)",
      borderLeft: "rgba(224, 157, 24, 0.9)",
    };
  };

  const alertIcon = (type) => (type === "dateSortieReelle" ? "✕" : "!");
  const alertIconStyle = (type) => {
    if (type === "dateSortieReelle") {
      return {
        background: "linear-gradient(180deg, #ef6a6a 0%, #ce3535 100%)",
        color: "#ffffff",
        boxShadow: "0 4px 10px rgba(206, 53, 53, 0.28)",
      };
    }
    return {
      background: "linear-gradient(180deg, #f3cf64 0%, #e39a33 100%)",
      color: "#7a3218",
      boxShadow: "0 4px 10px rgba(227, 154, 51, 0.24)",
    };
  };

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

      {/* Alerts (compact) */}
      {alerts.length > 0 && (
        <div style={{ ...card, padding: "8px 10px", marginBottom: 8 }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 5 }}>
            <span style={{ fontSize: 9, color: "#4a6170", letterSpacing: "0.08em", fontWeight: 700 }}>ALERTES</span>
            <span style={{ fontSize: 9, color: "#8e4d1e", fontWeight: 700 }}>{alerts.length}</span>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            {visibleAlerts.map((a) => {
              const s = alertStyle(a.type);
              const iconS = alertIconStyle(a.type);
              const person = persons.find((p) => p.id === a.personId);
              return (
                <button
                  key={a.key}
                  onClick={() => person && onSelectPerson(person)}
                  style={{ width: "100%", textAlign: "left", padding: "5px 8px", borderRadius: 8, border: `1px solid ${s.border}`, borderLeft: `3px solid ${s.borderLeft}`, background: s.bg, color: s.color, fontSize: 10, cursor: "pointer" }}
                  title={a.text}
                >
                  <span style={{ display: "flex", alignItems: "center", gap: 6, minWidth: 0 }}>
                    <span style={{ width: 18, height: 18, borderRadius: "50%", display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 11, fontWeight: 800, lineHeight: 1, flex: "0 0 auto", ...iconS }}>
                      {alertIcon(a.type)}
                    </span>
                    <span style={{ minWidth: 0, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{a.text}</span>
                  </span>
                </button>
              );
            })}
            {alerts.length > visibleAlerts.length && (
              <div style={{ fontSize: 9, color: "#556d79", textAlign: "right" }}>+{alerts.length - visibleAlerts.length} alerte(s)</div>
            )}
          </div>
        </div>
      )}

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
          const dossierStatus = getDossierStatus(p);
          const sc = statusColor(dossierStatus);
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
                {dossierStatus || "—"}
              </span>
            </button>
          );
        })}
      </div>
    </div>
  );
}
