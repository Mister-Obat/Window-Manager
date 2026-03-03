# Debug Tickets - Window Manager

## Règles
- Créer un ticket si un bug dépasse 1 session.
- Un ticket contient: contexte, symptômes, repro, hypothèses, actions, statut.
- Fermer un ticket après validation utilisateur.

## Template Ticket
- **ID**: `DBG-YYYYMMDD-XX`
- **Titre**:
- **Contexte**:
- **Symptômes**:
- **Étapes de reproduction**:
- **Résultat attendu**:
- **Résultat observé**:
- **Hypothèses**:
- **Actions menées**:
- **Décision**:
- **Statut**: `open | monitoring | resolved | closed`

## Tickets Actifs

- Aucun.

## Historique

### DBG-20260303-01
- **Titre**: Écran noir PowerTerminal après restauration Window Manager
- **Contexte**: PowerTerminal (Electron frameless) restauré au démarrage via scénario.
- **Symptômes**: fenêtre visible mais contenu noir.
- **Étapes de reproduction**:
1. Sauvegarder un scénario contenant `PowerTerminal`.
2. Restaurer le scénario (notamment au boot).
3. Observer PowerTerminal.
- **Résultat attendu**: UI Electron rendue normalement après repositionnement.
- **Résultat observé**: écran noir intermittent/persistant.
- **Hypothèses**:
1. Repaint forcé Win32 (`RedrawWindow`) inadapté à Electron.
2. "Resize jiggle" (+1px/-1px) force un état de rendu noir sur certaines fenêtres Chromium.
3. Relance brute `electron.exe . --no-sandbox` depuis `cmdline` sauvegardée peut démarrer PowerTerminal dans un chemin dev instable.
4. `SetWindowPos(...SWP_SHOWWINDOW)` + escalade `SW_RESTORE` dans la boucle de placement restent trop agressifs pour certaines fenêtres frameless Electron.
- **Actions menées**:
1. Analyse des appels de restore dans `wm_engine/restorer.py`.
2. Patch ciblé: skip `RedrawWindow + jiggle` pour `electron.exe`, `chrome.exe`, `msedge.exe`.
3. Vérification syntaxe Python (`py_compile`) OK.
4. Patch complémentaire: pour fenêtres Electron/Chromium, suppression de `SWP_SHOWWINDOW` et de l'escalade `SW_RESTORE` dans les retries de positionnement.
5. Patch complémentaire lancement: si `cmdline` ressemble à `electron.exe .` en cwd de projet, privilégier `start-dev.bat`, sinon `npm run dev`, sinon fallback commande brute.
- **Décision**: conserver les heuristiques agressives uniquement pour apps non Chromium, appliquer un chemin de lancement/placement plus sûr pour Electron.
- **Statut**: `closed` (validé utilisateur le 2026-03-03)
