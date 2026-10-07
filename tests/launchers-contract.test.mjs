import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";

const ROOT = process.cwd();
const EXPECTED_REMOTE = "https://nextboard-dev.github.io/Dotations/";
const CURRENT_LOCAL_PATH = "03.%20DOTATIONS/MODE%20HEBERGE%20-%20VERSION%20ACTIVE/index.html";
const OLD_DASHBOARD_PATTERN = /GESTION(?:%20| )DES(?:%20| )ACC|EFFETS(?:%20| )SENSIBLES|favicon-dark\.ico/i;

function read(file) {
  return fs.readFileSync(path.join(ROOT, file), "utf8");
}

function urlValue(source) {
  const match = source.match(/^URL=(.+)$/m);
  return match ? match[1].trim() : "";
}

test("lanceurs: les raccourcis PC et telephone ouvrent les bonnes entrees Dotations", () => {
  const pcHosted = read("Ouvrir-Dotations-PC-Heberge.url");
  const phoneHosted = read("Ouvrir-Dotations-Telephone-Heberge.url");
  const pcLocal = read("Ouvrir-Dotations-PC-Local.url");

  assert.equal(urlValue(pcHosted), `${EXPECTED_REMOTE}?view=desktop`);
  assert.equal(urlValue(phoneHosted), `${EXPECTED_REMOTE}?view=mobile`);
  assert.match(urlValue(pcLocal), /^file:\/\/\/C:\/Users\/sebastien\.duc\/CLOUD\/02_ARCHIVAGE%20PERSONNEL\/DASHBOARDS\//);
  assert.match(urlValue(pcLocal), new RegExp(CURRENT_LOCAL_PATH.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
});

test("lanceurs: aucun raccourci Dotations ne pointe vers un ancien dashboard ou ancien favicon", () => {
  const shortcuts = [
    "Ouvrir-Dotations-PC-Heberge.url",
    "Ouvrir-Dotations-PC-Local.url",
    "Ouvrir-Dotations-Telephone-Heberge.url",
  ];

  for (const shortcut of shortcuts) {
    assert.doesNotMatch(read(shortcut), OLD_DASHBOARD_PATTERN, shortcut);
  }
});

test("lanceur telephone local: il demarre le projet smartphone sur le reseau local", () => {
  const launcher = read("Ouvrir-Dotations-Telephone-Local.bat");

  assert.match(launcher, /set "APP_DIR=%ROOT%smartphone"/);
  assert.match(launcher, /if not exist "%APP_DIR%\\package\.json"/);
  assert.match(launcher, /where node >nul 2>nul/);
  assert.match(launcher, /where npm >nul 2>nul/);
  assert.match(launcher, /npm install/);
  assert.match(launcher, /--host 0\.0\.0\.0 --port 5173/);
  assert.match(launcher, /VITE_SMARTPHONE_READONLY=false/);
  assert.match(launcher, /http:\/\/%LOCAL_IP%:5173/);
});
