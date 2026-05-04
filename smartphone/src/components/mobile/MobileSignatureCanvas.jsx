import React, { useRef, useState, useEffect } from "react";
import { db } from "@/lib/db";

const MOBILE_SIGNATURE_CANVAS_CSS_HEIGHT = 220;

export default function MobileSignatureCanvas({ personId, docType, signer, signerLabel, existingSignature, onSaved, signataireName = "", signataireFunction = "", signataireId = "" }) {
  const canvasRef = useRef(null);
  const [drawing, setDrawing] = useState(false);
  const [signed, setSigned] = useState(false);
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState(null);
  const lastPos = useRef(null);
  const [loadedSignatureData, setLoadedSignatureData] = useState("");

  const setupCanvas = () => {
    const canvas = canvasRef.current;
    if (!canvas) return null;

    const rect = canvas.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    const cssWidth = Math.max(1, Math.round(rect.width));
    const cssHeight = Math.max(1, Math.round(rect.height));
    const needResize = canvas.width !== Math.round(cssWidth * dpr) || canvas.height !== Math.round(cssHeight * dpr);

    if (needResize) {
      canvas.width = Math.round(cssWidth * dpr);
      canvas.height = Math.round(cssHeight * dpr);
    }

    const ctx = canvas.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.lineWidth = 2.5;
    ctx.lineCap = "round";
    ctx.strokeStyle = "#1a2e38";
    return { canvas, ctx };
  };

  useEffect(() => {
    setupCanvas();
  }, []);

  useEffect(() => {
    const onResize = () => {
      const prepared = setupCanvas();
      if (!prepared) return;
      const source = loadedSignatureData || existingSignature?.signatureData;
      if (!source) return;
      const img = new Image();
      img.onload = () => {
        const { canvas, ctx } = prepared;
        ctx.clearRect(0, 0, canvas.clientWidth, canvas.clientHeight);
        ctx.drawImage(img, 0, 0, canvas.clientWidth, canvas.clientHeight);
        setSigned(true);
      };
      img.src = source;
    };
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, [existingSignature?.signatureData, loadedSignatureData]);

  useEffect(() => {
    const prepared = setupCanvas();
    if (!prepared) return;
    const { canvas, ctx } = prepared;
    const incoming = existingSignature?.signatureData || "";
    ctx.clearRect(0, 0, canvas.clientWidth, canvas.clientHeight);
    if (!incoming) {
      setLoadedSignatureData("");
      setSigned(false);
      return;
    }

    const img = new Image();
    img.onload = () => {
      ctx.clearRect(0, 0, canvas.clientWidth, canvas.clientHeight);
      ctx.drawImage(img, 0, 0, canvas.clientWidth, canvas.clientHeight);
      setLoadedSignatureData(incoming);
      setSigned(true);
    };
    img.src = incoming;
  }, [existingSignature]);

  const getPos = (e, canvas) => {
    const rect = canvas.getBoundingClientRect();
    if (e.touches) {
      return { x: e.touches[0].clientX - rect.left, y: e.touches[0].clientY - rect.top };
    }
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
  };

  const startDraw = (e) => {
    e.preventDefault();
    const prepared = setupCanvas();
    if (!prepared) return;
    const { canvas } = prepared;
    setDrawing(true);
    lastPos.current = getPos(e, canvas);
  };

  const draw = (e) => {
    if (!drawing) return;
    e.preventDefault();
    const prepared = setupCanvas();
    if (!prepared) return;
    const { canvas, ctx } = prepared;
    const pos = getPos(e, canvas);
    ctx.beginPath();
    ctx.moveTo(lastPos.current.x, lastPos.current.y);
    ctx.lineTo(pos.x, pos.y);
    ctx.stroke();
    lastPos.current = pos;
  };

  const endDraw = () => {
    setDrawing(false);
    setSigned(true);
  };

  const clear = () => {
    const prepared = setupCanvas();
    if (!prepared) return;
    const { canvas, ctx } = prepared;
    ctx.clearRect(0, 0, canvas.clientWidth, canvas.clientHeight);
    setSigned(false);
    setLoadedSignatureData("");
  };

  const save = async () => {
    const prepared = setupCanvas();
    if (!prepared) return;
    const { canvas } = prepared;
    const data = canvas.toDataURL("image/png");
    setSaving(true);
    try {
      const signedAt = new Date().toISOString();
      const existing = existingSignature;
      const payload = {
        signatureData: data,
        signedAt,
        signataireId: signataireId || null,
        signataireName: signataireName || null,
        signataireFunction: signataireFunction || null,
      };
      if (signer === "representant" && !payload.signataireName) {
        setMsg("CHOISIR LE REPRESENTANT");
        setTimeout(() => setMsg(null), 2500);
        return;
      }
      if (existing) {
        await db.Signature.update(existing.id, payload);
      } else {
        await db.Signature.create({ personId, docType, signer, ...payload });
      }
      setLoadedSignatureData(data);
      setSigned(true);
      setMsg("SIGNATURE VALIDEE ✓");
      setTimeout(() => setMsg(null), 2500);
      if (onSaved) {
        await onSaved({
          personId,
          docType,
          signer,
          signatureData: data,
          signedAt,
        });
      }
    } catch (error) {
      console.error("Signature save error:", error);
      const message = String(error?.message || "").toUpperCase();
      setMsg(message || "ERREUR DE SAUVEGARDE SUPABASE");
      setTimeout(() => setMsg(null), 2500);
    } finally {
      setSaving(false);
    }
  };

  return (
    <div style={{ background: "rgba(244,241,234,0.92)", border: "1px solid rgba(173,190,199,0.98)", borderRadius: 11, padding: "10px", marginBottom: 8 }}>
      <div style={{ fontSize: 10, color: "#3f5662", letterSpacing: "0.08em", marginBottom: 6, fontWeight: 600 }}>{signerLabel}</div>

      <div style={{ position: "relative", background: "#fffdfa", border: "1px dashed rgba(63,97,112,0.4)", borderRadius: 9, overflow: "hidden", touchAction: "none" }}>
        <canvas
          ref={canvasRef}
          width={640}
          height={280}
          style={{ display: "block", width: "100%", height: MOBILE_SIGNATURE_CANVAS_CSS_HEIGHT, cursor: "crosshair", touchAction: "none" }}
          onMouseDown={startDraw}
          onMouseMove={draw}
          onMouseUp={endDraw}
          onMouseLeave={endDraw}
          onTouchStart={startDraw}
          onTouchMove={draw}
          onTouchEnd={endDraw}
        />
        {!signed && (
          <div style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", pointerEvents: "none" }}>
            <span style={{ fontSize: 11, color: "rgba(63,97,112,0.4)", letterSpacing: "0.06em" }}>SIGNER ICI</span>
          </div>
        )}
      </div>

      {msg && <div style={{ margin: "6px 0", fontSize: 10, color: "#2e6a44", fontWeight: 600 }}>{msg}</div>}

      <div style={{ display: "flex", gap: 6, marginTop: 8 }}>
        <button onClick={save} disabled={!signed || saving} style={{ flex: 1, padding: "8px 10px", borderRadius: 9, border: "none", background: signed ? "#3f6170" : "rgba(63,97,112,0.3)", color: "#fff", fontSize: 10, fontWeight: 700, cursor: signed ? "pointer" : "default" }}>
          {saving ? "..." : "VALIDER"}
        </button>
        <button onClick={clear} style={{ flex: 1, padding: "8px 10px", borderRadius: 9, border: "1px solid rgba(63,97,112,0.3)", background: "rgba(63,97,112,0.1)", color: "#213b48", fontSize: 10, cursor: "pointer" }}>
          EFFACER
        </button>
      </div>
    </div>
  );
}
