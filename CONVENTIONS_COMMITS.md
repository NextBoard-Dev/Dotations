# CONVENTIONS DE COMMITS (DOTATIONS)

Objectif: ne plus produire de commits avec titres techniques du type `Reprise C:\...`.

## Regle simple

- Format obligatoire: `type(scope): resume court`

## Types autorises

- `fix`: correction
- `feat`: nouvelle fonctionnalite
- `docs`: documentation
- `chore`: menage technique

## Exemples valides

- `fix(effect-form): garder designation active quel que soit le contexte`
- `fix(sticky): supprimer le fond visible derriere les zones top-fixed`
- `feat(ui-motion): animer KPI, sauvegarde et validation signature`
- `docs(project): mettre a jour les regles du dossier Dotations`
- `chore(repo): normaliser la convention des messages de commit`

## Interdits

- Messages generiques sans contexte (`update`, `modifs`, `test`)
- Messages auto contenant un chemin local Windows

