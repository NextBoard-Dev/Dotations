import React from "react";
import { getEffectStatus } from "@/lib/businessRules";

const STATUT_COLORS = {
  "ACTIF": { bg: "rgba(89,148,117,0.16)", color: "#2f5e43" },
  "RESTITUE": { bg: "rgba(87,143,106,0.2)", color: "#2c513a" },
  "NON RENDU": { bg: "rgba(224,147,82,0.2)", color: "#8e4d1e" },
  "PERDU": { bg: "rgba(202,91,96,0.19)", color: "#7d2a31" },
  "HS": { bg: "rgba(132,140,149,0.22)", color: "#4a545d" },
  "VOL": { bg: "rgba(181,120,172,0.2)", color: "#6f3d73" },
  "VOLE": { bg: "rgba(181,120,172,0.2)", color: "#6f3d73" },
  "DETRUIT": { bg: "rgba(122,112,170,0.19)", color: "#4f447f" },
};

export default function MobileEffetsList({ effets, onEdit }) {
  const formatCost = (value) => {
    const amount = Number(value);
    return `${Number.isFinite(amount) ? amount : 0}€`;
  };
  const hasPositiveCost = (value) => {
    const amount = Number(value);
    return Number.isFinite(amount) && amount > 0;
  };

  if (effets.length === 0) {
    return (
      <div style={{ background: "rgba(244,241,234,0.98)", border: "1px solid rgba(173,190,199,0.98)", borderRadius: 11, padding: 16, textAlign: "center", color: "#3f5662", fontSize: 11 }}>
        AUCUN EFFET CONFIE
      </div>
    );
  }

  return (
    <div>
      {effets.map(e => {
        const displayStatus = getEffectStatus(null, e);
        const sc = STATUT_COLORS[displayStatus] || STATUT_COLORS["ACTIF"];
        return (
          <button
            key={e.id}
            onClick={() => onEdit(e)}
            style={{ display: "block", width: "100%", textAlign: "left", background: "rgba(244,241,234,0.98)", border: "1px solid rgba(173,190,199,0.98)", borderRadius: 11, padding: "10px 12px", marginBottom: 6, boxShadow: "0 2px 8px rgba(31,49,59,0.07)", cursor: "pointer" }}
          >
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 4 }}>
              <span style={{ fontWeight: 700, fontSize: 12, color: "#0f1e26" }}>{e.typeEffet}</span>
              <span style={{ fontSize: 9, padding: "2px 8px", borderRadius: 99, background: sc.bg, color: sc.color }}>{displayStatus}</span>
            </div>
            <div style={{ fontSize: 11, color: "#213b48", marginBottom: 2 }}>{e.designation || "—"}</div>
            <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
              {e.numeroIdentification && <span style={{ fontSize: 9, color: "#3f5662" }}>N° {e.numeroIdentification}</span>}
              {e.dateRemise && <span style={{ fontSize: 9, color: "#3f5662" }}>Remis le {e.dateRemise}</span>}
              {hasPositiveCost(e.coutRemplacement) && (
                <span style={{ fontSize: 9, color: "#9b5a2a", fontWeight: 600 }}>{formatCost(e.coutRemplacement)}</span>
              )}
            </div>
            {e.commentaire && <div style={{ fontSize: 9, color: "#556d79", marginTop: 3, fontStyle: "italic" }}>{e.commentaire}</div>}
          </button>
        );
      })}
    </div>
  );
}
