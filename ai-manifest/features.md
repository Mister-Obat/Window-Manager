# Features Details: Window Manager

## 1. Sauvegarde de scénarios
- Capture des fenêtres éligibles.
- Enregistrement position (`rect`), état (`show_cmd`) et contexte process.
- Support de noms de scénarios configurables (slots UI).

## 2. Restauration intelligente
- Phase 1: tentative de placement des fenêtres déjà ouvertes.
- Phase 2: lancement des fenêtres manquantes puis placement.
- Gestion dédiée navigateurs (Chrome/Firefox/Edge) et mode privé.

## 3. Réglages
- `restore_minimized`, `precise_urls`, filtres d'inclusion/exclusion.
- Overrides par scénario pour adapter la stratégie de scan/restore.

## 4. Correctif récent (2026-03-03)
- Problème: `PowerTerminal` (Electron) pouvait afficher une fenêtre noire après restauration.
- Causes:
1. Combinaison `SW_RESTORE` + `RedrawWindow` + "resize jiggle" (`w+1/w`) appliquée à des fenêtres Chromium/Electron.
2. Relance brute de commandes dev Electron (`electron.exe . --no-sandbox`) depuis `cmdline` sauvegardée.
- Corrections dans `wm_engine/restorer.py`:
1. Bloc repaint+jiggle ignoré pour `electron.exe`, `chrome.exe`, `msedge.exe`.
2. Positionnement Electron/Chromium adouci: pas de `SWP_SHOWWINDOW`, pas d'escalade `SW_RESTORE` dans les retries.
3. Relance Electron dev durcie: priorité `start-dev.bat`, sinon `npm run dev`, puis fallback commande brute.

## Limites connues
- Le lancement d'apps en mode dev reste dépendant de l'environnement local (scripts/outils manquants => fallback commande brute potentiellement instable).
- Le matching reste heuristique pour certains titres dynamiques.
