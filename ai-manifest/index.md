# Window Manager Manifest

## Vision
Window Manager sauvegarde puis restaure des dispositions de fenêtres Windows (position, taille, état, contexte) de façon fiable après redémarrage ou relance d'apps.

## Scope Actuel

### 1. Capture de layout
- Énumération des fenêtres visibles via Win32.
- Filtrage par titre/classe/settings.
- Sauvegarde des métadonnées utiles: `rect`, `show_cmd`, `cmdline`, `cwd`, URL navigateur, dossier Explorer.

### 2. Restauration de layout
- Matching des fenêtres existantes puis placement immédiat.
- Lancement séquentiel des fenêtres manquantes puis placement.
- Support états normal / minimisé / maximisé.

### 3. Configuration utilisateur
- Réglages globaux + overrides par slot/scénario.
- Exclusions de titres/processus.
- Option démarrage automatique de l'app au boot.

## Invariants & Contraintes
- Correctness first: une fenêtre ne doit pas être matchée à la mauvaise cible.
- Le placement doit rester déterministe, même avec fenêtres minimisées.
- Les heuristiques agressives de repaint/résize doivent rester sûres pour les apps modernes (Electron/Chromium).
- Toute modif de logique de restore doit être traçable dans `ai-manifest/debug/README.md`.

## Structure /ai-manifest/
- `index.md`: vision et invariants globaux.
- `tech-stack.md`: stack et APIs système.
- `features.md`: fonctionnalités et limites.
- `design.md`: principes UI et UX.
- `user-flows.md`: parcours utilisateur.
- `debug/README.md`: tickets actifs et historique debug.
