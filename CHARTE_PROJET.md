# CHARTE PROJET (DOTATIONS)

Date de mise a jour: 16/03/2026

## 1) SOURCE MAITRE

- Le dossier maitre de travail est `Dotations`.
- Les modifications se font uniquement dans `Dotations`.
- Aucune edition dans une copie parallele.

## 2) SNAPSHOTS

- Les snapshots sont autorises uniquement hors repo.
- Interdiction de creer ou commiter `snapshots/` dans `Dotations`.
- Si besoin: utiliser un dossier externe de sauvegarde locale.

## 3) REGLES UI METIER (ACTUELLES)

- Carte Effet: champs contextuels grises/inactifs, non masques.
- Les libelles changent selon le type d'effet.
- Liaison base de reference active:
  - site dynamique par type,
  - nom existant filtre par type + site.
- `STATUT MANUEL`, `TYPE D'EFFET` et `SITE` (quand requis) sont obligatoires.
- `N° telecommande / N° carte` = conseille, non obligatoire.

## 4) DOCUMENTS ET SIGNATURES

- Un document n'entre dans `Documents archives` que s'il est signe des 2 cotes.
- Bouton `OUVRIR EN PDF` bloque tant que signatures non completes.
- Alerte explicite si tentative d'ouverture PDF sans double signature.

## 5) QUALITE VISUELLE

- Les parties hautes non scrollables doivent rester stables au scroll.
- Les 5 boutons de la carte Effet restent sur une seule ligne.
- Les KPI et couts doivent rester coherents avec les statuts.

## 6) COMMITS

- Format impose: `type(scope): resume court`
- Exemples:
  - `fix(sticky): stabiliser les blocs top-fixed`
  - `fix(effect-form): champs contextuels grises`
  - `feat(ui-motion): animations sobres formulaire`

