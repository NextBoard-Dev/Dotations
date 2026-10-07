import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";

const checkRepo = fs.readFileSync("check_repo.bat", "utf8");
const gitignore = fs.readFileSync(".gitignore", "utf8");

test("controle depot: check_repo valide le bon remote et la branche courante", () => {
  assert.match(checkRepo, /git rev-parse --show-toplevel/);
  assert.match(checkRepo, /git branch --show-current/);
  assert.match(checkRepo, /git remote get-url origin/);
  assert.match(checkRepo, /https:\/\/github\.com\/NextBoard-Dev\/Dotations\.git/);
  assert.match(checkRepo, /REMOTE INATTENDU POUR DOTATIONS/);
});

test("controle depot: check_repo lance les controles critiques complets", () => {
  assert.match(checkRepo, /node --check app\.js \|\| goto :fail/);
  assert.match(checkRepo, /node --test tests\/\*\.mjs \|\| goto :fail/);
  assert.match(checkRepo, /\[ECHEC\] CONTROLES CRITIQUES DOTATIONS/);
  assert.match(checkRepo, /OK - CONTROLES CRITIQUES VALIDES/);
});

test("controle depot: les sauvegardes locales restent ignorees", () => {
  assert.match(gitignore, /snapshots\//);
  assert.match(gitignore, /snapshots_ok\//);
  assert.match(gitignore, /LATEST_MAJOR_SNAPSHOT\.txt/);
  assert.match(gitignore, /sauvegarde Github\//);
  assert.match(gitignore, /sauvegarde Supabase\//);
});
