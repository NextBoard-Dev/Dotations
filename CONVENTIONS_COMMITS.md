# CONVENTIONS DE COMMITS (DOTATIONS)

Date de mise a jour: 18/03/2026

## Regle simple

- Phrase de commit courte et lisible.
- Format recommande: `type(scope): resume court`.

## Types autorises

- `fix`: correction
- `feat`: nouvelle fonctionnalite
- `docs`: documentation
- `chore`: maintenance

## Exemples valides

- `fix(pdf): animation fleche verte quand document completement signe`
- `fix(effect-form): multi-selection stable sans casser les calculs`
- `docs(project): mise a jour des regles Dotations`

## Regles de session

- Apres chaque correction: proposer une phrase de commit.
- Si l'utilisateur ecrit `go commit`: commit local automatique.
- Reponse post-commit obligatoire: `COMMIT TERMINE - TU PEUX PUSH`.
- Le push reste manuel (GitHub Desktop par l'utilisateur).

## Interdits

- Messages vagues (`update`, `test`, `modifs`).
- Messages auto avec chemin local Windows.
- Commits melangeant des changements hors besoin.
