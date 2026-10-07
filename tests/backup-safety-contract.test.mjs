import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";

const backupScript = fs.readFileSync("scripts/backup_dotations_edge.ps1", "utf8");
const schedulerScript = fs.readFileSync("scripts/planifier_backup_quotidien.bat", "utf8");
const gitignore = fs.readFileSync(".gitignore", "utf8");

test("backup Supabase: le script utilise une authentification utilisateur et une cle publiable", () => {
  assert.match(backupScript, /\[Parameter\(Mandatory = \$true\)\]\s*\r?\n\s*\[string\]\$Email/);
  assert.match(backupScript, /\[Parameter\(Mandatory = \$true\)\]\s*\r?\n\s*\[string\]\$Password/);
  assert.match(backupScript, /\$publishableKey = "sb_publishable_/);
  assert.doesNotMatch(backupScript, /service[_-]?role/i);
  assert.match(backupScript, /auth\/v1\/token\?grant_type=password/);
  assert.match(backupScript, /Authorization = "Bearer \$token"/);
});

test("backup Supabase: l'export passe par l'Edge Function et conserve un fichier JSON", () => {
  assert.match(backupScript, /functions\/v1\/dotations-api\/data/);
  assert.match(backupScript, /DOTATIONS SNAPSHOTS/);
  assert.match(backupScript, /backup_dotations_/);
  assert.match(backupScript, /Set-Content -LiteralPath \$outFile -Encoding UTF8/);
  assert.match(backupScript, /BACKUP_OK:\$outFile/);
});

test("backup planifie: le lanceur reste portable et pointe vers le script voisin", () => {
  assert.match(schedulerScript, /set "SCRIPT_PS=%~dp0backup_dotations_edge\.ps1"/);
  assert.doesNotMatch(schedulerScript, /C:\\Users\\sebastien\.duc\\CLOUD\\02_ARCHIVAGE PERSONNEL\\DASHBOARDS\\DOTATIONS\\/i);
  assert.match(schedulerScript, /if not exist "%SCRIPT_PS%"/);
  assert.match(schedulerScript, /schtasks \/Create \/F \/SC DAILY \/ST 03:30/);
  assert.match(schedulerScript, /-Email \\"%EMAIL%\\" -Password \\"%PASS%\\"/);
});

test("backup: les sorties locales de sauvegarde restent ignorees par Git", () => {
  assert.match(gitignore, /snapshots\//);
  assert.match(gitignore, /sauvegarde Supabase\//);
  assert.match(gitignore, /supabase\/\.temp\//);
});
