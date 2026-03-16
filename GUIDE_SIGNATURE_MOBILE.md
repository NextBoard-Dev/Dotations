# GUIDE SIGNATURE MOBILE (DOTATIONS)

Date de mise a jour: 16/03/2026

## Principe

- Chaque document a 2 signatures distinctes:
  - Personnel
  - Representant
- La validation doit enregistrer image + date de validation.

## Regle PDF/Archives

- `OUVRIR EN PDF` doit etre bloque tant que les 2 signatures ne sont pas valides.
- Un document non doublement signe ne doit pas apparaitre dans `Documents archives`.

## Mode heberge

- URL publique configuree dans les parametres.
- QR doit ouvrir la bonne cible de signature.

## Verifications minimales

1. Scanner QR Personnel -> signature Personnel.
2. Scanner QR Representant -> signature Representant.
3. Valider les 2 signatures.
4. Verifier activation PDF.
5. Verifier apparition dans Archives.

