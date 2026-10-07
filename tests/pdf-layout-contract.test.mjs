import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";

const appJs = fs.readFileSync("app.js", "utf8");
const styleCss = fs.readFileSync("style.css", "utf8");
const arrivalHtml = fs.readFileSync("document-arrivee.html", "utf8");
const exitHtml = fs.readFileSync("document-sortie.html", "utf8");

function normalize(source) {
  return source.replace(/\s+/g, " ").trim();
}

function getTableBlock(html, bodyId) {
  const bodyIndex = html.indexOf(`id="${bodyId}"`);
  assert.notEqual(bodyIndex, -1, `table body ${bodyId} absent`);
  const tableStart = html.lastIndexOf("<table", bodyIndex);
  const tableEnd = html.indexOf("</table>", bodyIndex);
  assert.notEqual(tableStart, -1, `table start for ${bodyId} absent`);
  assert.notEqual(tableEnd, -1, `table end for ${bodyId} absent`);
  return html.slice(tableStart, tableEnd);
}

function countHeaders(tableBlock) {
  return (tableBlock.match(/<th\b/g) || []).length;
}

function countSignatureCanvases(html, docType) {
  const matches = html.match(new RegExp(`class="signature-box__canvas js-signature-canvas"[\\s\\S]*?data-doc-type="${docType}"`, "g")) || [];
  return matches.length;
}

test("PDF: le cache est invalidé quand le contrat de mise en page change", () => {
  assert.match(appJs, /const PDF_LAYOUT_VERSION = "2026-10-07-exit-table-compact-v1";/);
  assert.match(appJs, /layoutVersion: PDF_LAYOUT_VERSION/);
  assert.match(appJs, /const PDF_FORMAT_LOCK = "v1";/);
  assert.match(appJs, /document\.body\.dataset\.pdfLayoutLock = PDF_FORMAT_LOCK/);
  assert.match(styleCss, /\[data-pdf-layout-lock="v1"\]/);
});

test("PDF arrivée: structure document, signatures et colonnes figées", () => {
  const table = getTableBlock(arrivalHtml, "arrival-effects-body");

  assert.equal(countHeaders(table), 7);
  assert.match(table, /colspan="7" class="table-empty"/);
  assert.equal(countSignatureCanvases(arrivalHtml, "arrival"), 2);
  assert.match(arrivalHtml, /data-signer="personnel"/);
  assert.match(arrivalHtml, /data-signer="representant"/);
  assert.match(appJs, /buildEmptyTableRow\(body, "AUCUN EFFET A AFFICHER", 7\)/);
  assert.match(appJs, /colspan="\$\{isPdfMode \? "5" : "6"\}"/);
});

test("PDF sortie: structure document, signatures et colonnes figées", () => {
  const table = getTableBlock(exitHtml, "exit-effects-body");

  assert.equal(countHeaders(table), 11);
  assert.match(table, /RENDU CE JOUR/);
  assert.match(table, /COUT FACTURABLE/);
  assert.match(table, /FACTURATION/);
  assert.match(table, /colspan="11" class="table-empty"/);
  assert.equal(countSignatureCanvases(exitHtml, "exit"), 2);
  assert.match(exitHtml, /data-signer="personnel"/);
  assert.match(exitHtml, /data-signer="representant"/);
  assert.match(appJs, /buildEmptyTableRow\(body, "AUCUN EFFET A AFFICHER", 11\)/);
  assert.match(appJs, /colspan="\$\{isPdfMode \? "7" : "10"\}"/);
});

test("PDF sortie: affichage compact sans trou visuel dans le tableau", () => {
  const css = normalize(styleCss);

  assert.match(css, /#exit-effects-body tr .*height: auto !important; min-height: 0 !important;/);
  assert.match(css, /#exit-effects-body td .*height: auto !important; min-height: 0 !important;/);
  assert.match(css, /data-page="exit-document".*grid-template-columns: 15% 17% 13% 14% 14% 10% 9% 8%;/);
  assert.match(css, /nth-child\(8\).*nth-child\(10\).*nth-child\(11\).*display: none;/);
  assert.doesNotMatch(css, /nth-child\(12\).*display: none;/);
  assert.match(css, /table-total-row \{ grid-template-columns: 1fr auto; \}/);
  assert.match(css, /table-total-row td:last-child .*grid-column: 2; min-width: 54px;/);
});

test("PDF: les controles interactifs sont exclus du rendu imprime", () => {
  const css = normalize(styleCss);

  assert.match(css, /body\[data-pdf-mode="true"\] \.sidebar,/);
  assert.match(css, /body\[data-pdf-mode="true"\] \.document-picker-section \{ display: none !important; \}/);
  assert.match(css, /body\[data-pdf-mode="true"\] \.signature-box__actions \{ display: none !important; \}/);
  assert.match(css, /body\[data-pdf-mode="true"\] \.signature-box__mobile-link \{ display: none !important; \}/);
  assert.match(css, /body\[data-pdf-mode="true"\] \.document-effects-action-col \{ display: none; \}/);
  assert.match(css, /@page \{ size: A4; margin: 5mm 7mm; \}/);
});
