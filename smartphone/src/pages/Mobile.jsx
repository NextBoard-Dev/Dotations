import React, { useState, useEffect, useRef } from "react";
import MobileOverview from "../components/mobile/MobileOverview";
import MobileFichePerson from "../components/mobile/MobileFichePerson";
import MobileDocumentArrivee from "../components/mobile/MobileDocumentArrivee";
import MobileDocumentSortie from "../components/mobile/MobileDocumentSortie";
import { db } from "@/lib/db";

const TABS = [
  { id: "overview", label: "VUE D'ENSEMBLE", icon: "🏠" },
  { id: "fiche", label: "FICHE", icon: "👤" },
  { id: "arrivee", label: "ENTREE", icon: "📥" },
  { id: "sortie", label: "SORTIE", icon: "📤" },
];

const DEFAULT_BASES = {
  sites: [],
  fonctions: [],
  typesPersonnel: [],
  typesContrats: [],
  typesEffets: [],
  statutsObjetManuels: [],
  representantsSignataires: [],
};

function isValidTab(tab) {
  return TABS.some((t) => t.id === tab);
}

function buildUrlState(tab, personId) {
  const params = new URLSearchParams();
  params.set("tab", isValidTab(tab) ? tab : "overview");
  if (personId) params.set("personId", String(personId));
  const query = params.toString();
  return `${window.location.pathname}${query ? `?${query}` : ""}`;
}

function readUrlState() {
  const params = new URLSearchParams(window.location.search);
  const tab = params.get("tab");
  const personId = params.get("personId");
  return {
    tab: isValidTab(tab) ? tab : "overview",
    personId: personId || null,
  };
}

export default function Mobile() {
  const initialUrl = readUrlState();
  const [activeTab, setActiveTab] = useState(initialUrl.tab);
  const [persons, setPersons] = useState([]);
  const [effets, setEffets] = useState([]);
  const [bases, setBases] = useState(DEFAULT_BASES);
  const [selectedPerson, setSelectedPerson] = useState(null);
  const [loading, setLoading] = useState(true);
  const [saveStatus, setSaveStatus] = useState("saved");

  const personsRef = useRef([]);
  const selectedPersonRef = useRef(null);
  const skipNextPushRef = useRef(true);

  const applyUrlState = (state) => {
    const wantedTab = isValidTab(state?.tab) ? state.tab : "overview";
    const wantedPersonId = state?.personId || null;
    setActiveTab(wantedTab);

    if (!wantedPersonId) {
      setSelectedPerson(null);
      return;
    }

    const found = personsRef.current.find((p) => String(p.id) === String(wantedPersonId));
    setSelectedPerson(found || null);
  };

  const loadData = async () => {
    setLoading(true);
    try {
      const [pRes, eRes, bRes] = await Promise.allSettled([
        db.Person.list("-created_at", 5000),
        db.Effet.list("-created_at", 5000),
        db.AppState.getReferenceBases(),
      ]);

      const p = pRes.status === "fulfilled" && Array.isArray(pRes.value) ? pRes.value : [];
      const e = eRes.status === "fulfilled" && Array.isArray(eRes.value) ? eRes.value : [];
      const b = bRes.status === "fulfilled" && bRes.value ? bRes.value : DEFAULT_BASES;

      personsRef.current = p;
      setPersons(p);
      setEffets(e);
      setBases(b);

      const urlState = readUrlState();
      if (urlState.personId) {
        const found = p.find((person) => String(person.id) === String(urlState.personId));
        setSelectedPerson(found || null);
      }
      if (!isValidTab(urlState.tab)) {
        setActiveTab("overview");
      }
    } catch {
      personsRef.current = [];
      setPersons([]);
      setEffets([]);
      setBases(DEFAULT_BASES);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadData();
  }, []);

  useEffect(() => {
    selectedPersonRef.current = selectedPerson;
  }, [selectedPerson]);

  useEffect(() => {
    const onPopState = () => {
      skipNextPushRef.current = true;
      applyUrlState(readUrlState());
    };

    window.addEventListener("popstate", onPopState);
    return () => window.removeEventListener("popstate", onPopState);
  }, []);

  useEffect(() => {
    if (loading) return;

    if (skipNextPushRef.current) {
      skipNextPushRef.current = false;
      const replaceUrl = buildUrlState(activeTab, selectedPerson?.id);
      window.history.replaceState({}, "", replaceUrl);
      return;
    }

    const nextUrl = buildUrlState(activeTab, selectedPerson?.id);
    const currentUrl = `${window.location.pathname}${window.location.search}`;
    if (nextUrl !== currentUrl) {
      window.history.pushState({}, "", nextUrl);
    }
  }, [activeTab, selectedPerson?.id, loading]);

  const markUnsaved = () => setSaveStatus("unsaved");

  const handleSelectPerson = (person) => {
    setSelectedPerson(person);
  };

  const handleNavigateTo = (tab, person) => {
    if (person) setSelectedPerson(person);
    setActiveTab(tab);
  };

  const topSaveButtonStyle = (() => {
    if (saveStatus === "unsaved") {
      return {
        background: "#163b70",
        color: "#ffffff",
        border: "1px solid #0f2f59",
      };
    }
    if (saveStatus === "saving") {
      return {
        background: "rgba(63,97,112,0.24)",
        color: "#213b48",
        border: "1px solid rgba(63,97,112,0.35)",
      };
    }
    return {
      background: "rgba(111,157,120,0.2)",
      color: "#4c6a53",
      border: "1px solid rgba(111,157,120,0.3)",
    };
  })();

  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(180deg, #ebe6dc 0%, #d9e2e7 100%)", display: "flex", flexDirection: "column", maxWidth: 480, margin: "0 auto", position: "relative" }}>
      <div style={{ background: "linear-gradient(180deg, #c2d2da 0%, #d9e2e7 100%)", padding: "10px 14px 8px", borderBottom: "1px solid rgba(63,97,112,0.2)", display: "flex", alignItems: "center", gap: 10 }}>
        <img src="https://dphrvdhqhgycmllietuk.supabase.co/storage/v1/object/public/ui-assets/sidebar/bandeau-nextboard-sidebar-detoure.png" alt="NextBoard" style={{ height: 32, borderRadius: 6 }} />
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: 9, color: "#556d79", letterSpacing: "0.12em" }}>SUIVI DES DOTATIONS</div>
          <div style={{ fontSize: 11, fontWeight: 700, color: "#14242c", lineHeight: 1.2 }}>ENTREE / SORTIE</div>
        </div>
        <div style={{ display: "flex", gap: 5, alignItems: "center" }}>
          <button
            type="button"
            style={{
              fontSize: 8,
              padding: "0 9px",
              height: 26,
              borderRadius: 7,
              fontWeight: 700,
              letterSpacing: "0.04em",
              cursor: "default",
              ...topSaveButtonStyle,
            }}
          >
            {saveStatus === "saving" ? "SAUVEGARDE..." : saveStatus === "unsaved" ? "SAUVEGARDER" : "SAUVEGARDE"}
          </button>
          <button onClick={loadData} style={{ fontSize: 9, padding: "0 8px", height: 26, borderRadius: 7, border: "1px solid rgba(63,97,112,0.3)", background: "rgba(63,97,112,0.12)", color: "#213b48", cursor: "pointer" }}>
            ↻
          </button>
        </div>
      </div>

      <div style={{ flex: 1, overflowY: "auto", paddingBottom: 70 }}>
        {loading ? (
          <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: 200, color: "#3f5662", fontSize: 12 }}>
            CHARGEMENT...
          </div>
        ) : (
          <>
            {activeTab === "overview" && (
              <MobileOverview persons={persons} effets={effets} onSelectPerson={(p) => handleNavigateTo("fiche", p)} />
            )}
            {activeTab === "fiche" && (
              <MobileFichePerson
                persons={persons}
                effets={effets}
                selectedPerson={selectedPerson}
                onSelectPerson={handleSelectPerson}
                onDataChange={loadData}
                onMarkUnsaved={markUnsaved}
                setSaveStatus={setSaveStatus}
                onNavigate={handleNavigateTo}
                bases={bases}
              />
            )}
            {activeTab === "arrivee" && (
              <MobileDocumentArrivee
                persons={persons}
                effets={effets}
                selectedPerson={selectedPerson}
                onSelectPerson={handleSelectPerson}
                setSaveStatus={setSaveStatus}
                onDataChange={loadData}
                representatives={bases.representantsSignataires || []}
                pricingRules={bases.coutsRemplacement || []}
                effetTypes={bases.typesEffets || []}
              />
            )}
            {activeTab === "sortie" && (
              <MobileDocumentSortie
                persons={persons}
                effets={effets}
                selectedPerson={selectedPerson}
                onSelectPerson={handleSelectPerson}
                setSaveStatus={setSaveStatus}
                onDataChange={loadData}
                representatives={bases.representantsSignataires || []}
                pricingRules={bases.coutsRemplacement || []}
                effetTypes={bases.typesEffets || []}
              />
            )}
          </>
        )}
      </div>

      <div style={{ position: "fixed", bottom: 0, left: "50%", transform: "translateX(-50%)", width: "100%", maxWidth: 480, background: "linear-gradient(180deg, #c2d2da 0%, #b8cad2 100%)", borderTop: "1px solid rgba(63,97,112,0.25)", display: "grid", gridTemplateColumns: `repeat(${TABS.length}, 1fr)`, zIndex: 100 }}>
        {TABS.map((tab) => (
          <button
            key={tab.id}
            onClick={() => setActiveTab(tab.id)}
            style={{
              padding: "8px 4px 10px",
              border: "none",
              background: activeTab === tab.id ? "rgba(63,97,112,0.22)" : "transparent",
              borderTop: activeTab === tab.id ? "2px solid #3f6170" : "2px solid transparent",
              cursor: "pointer",
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              gap: 2,
            }}
          >
            <span style={{ fontSize: 18 }}>{tab.icon}</span>
            <span style={{ fontSize: 8, letterSpacing: "0.06em", color: activeTab === tab.id ? "#213b48" : "#556d79", fontWeight: activeTab === tab.id ? 700 : 400 }}>{tab.label}</span>
          </button>
        ))}
      </div>
    </div>
  );
}
