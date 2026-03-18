# DEPLOIEMENT GITHUB PAGES + SUPABASE (DOTATIONS)

Date de mise a jour: 18/03/2026

## 1) Pre-requis

- Repo GitHub: `Mililumatt/Dotations`
- Branche: `main`
- Source Pages: `main / root`

## 2) Checklist avant commit/push

- Aucune modification hors `Dotations`.
- Aucun snapshot dans le repo.
- Syntaxe JS verifiee (`node --check app.js`).
- Regles metier critiques valides:
  - blocage PDF sans double signature,
  - archives uniquement documents signes,
  - calculs et tableaux non casses.

## 3) Publication

1. Commit local.
2. Push `origin/main` (fait par l'utilisateur).
3. Verifier l'URL publique GitHub Pages.

## 4) Verification post-deploiement

- Documents arrivee/sortie:
  - signatures personnel + representant,
  - activation visuelle du bouton PDF quand 2 signatures valides,
  - ouverture PDF fonctionnelle.
- Archives:
  - presence uniquement des documents valides.
