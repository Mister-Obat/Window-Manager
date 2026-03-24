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

### DBG-20260324-01
- **Titre**: Une app introuvable bloque le flux de restauration séquentielle
- **Contexte**: scénario contenant au moins une entrée dont l'exécutable, le dossier ou la commande n'existe plus localement.
- **Symptômes**: après un échec de lancement, la restauration attend inutilement la fenêtre absente; les apps suivantes sont retardées et certaines fenêtres démarrées tard ne reçoivent pas leur placement final.
- **Étapes de reproduction**:
1. Sauvegarder un scénario contenant plusieurs apps.
2. Rendre introuvable l'une des cibles (exe supprimé, dossier déplacé, script absent).
3. Restaurer le scénario.
- **Résultat attendu**: l'entrée en erreur est journalisée puis ignorée sans bloquer la suite; les fenêtres qui apparaissent tardivement reçoivent encore leur placement.
- **Résultat observé**: attente prolongée sur l'item cassé puis absence de reprise finale pour certaines fenêtres restantes.
- **Hypothèses**:
1. `_launch_app` et `_launch_browser_group` journalisent l'erreur mais ne remontent pas l'échec à `restore_layout`.
2. La phase 2 attend donc quand même `_wait_for_window` sur une fenêtre qui ne viendra jamais.
3. L'orchestrateur n'a pas de passe finale mutualisée pour replacer les fenêtres apparues après leur fenêtre d'attente initiale.
4. Certains `cwd` sauvegardés sont versionnés (`Chrome`, `Discord`) et deviennent invalides après mise à jour, ce qui casse `subprocess.Popen(...)` malgré un exécutable principal encore disponible.
- **Actions menées**:
1. Durcissement de `wm_engine/restorer.py` pour retourner explicitement le résultat des lancements.
2. Skip immédiat de l'attente quand le lancement a échoué.
3. Ajout d'une phase 3 de placement final pour les fenêtres lancées mais détectées tardivement.
4. Suppression de code mort et d'accès dictionnaire fragiles dans le chemin de restauration.
5. Ajout d'un fallback de lancement: si le `cwd` sauvegardé est invalide, utiliser le dossier de l'exécutable, sinon lancer sans `cwd`.
- **Décision**: privilégier une orchestration tolérante aux erreurs partielles, sans abandon global de la restauration.
- **Statut**: `monitoring`

### DBG-20260324-02
- **Titre**: Changement d'ordre d'écrans ou d'écran principal casse le restore
- **Contexte**: scénario multi-écrans sauvegardé, puis modification de la topologie Windows (ordre des écrans, écran principal, position relative).
- **Symptômes**: les fenêtres sont replacées sur de mauvais écrans ou avec de mauvais offsets alors qu'elles restent techniquement "on-screen".
- **Étapes de reproduction**:
1. Sauvegarder un scénario avec plusieurs fenêtres réparties sur plusieurs écrans.
2. Modifier l'ordre des écrans ou changer l'écran principal dans Windows.
3. Restaurer le scénario.
- **Résultat attendu**: les fenêtres retrouvent un placement cohérent sur leurs écrans logiques malgré le changement de topologie.
- **Résultat observé**: les coordonnées absolues sauvegardées ne correspondent plus à la nouvelle topologie et le placement devient incorrect.
- **Hypothèses**:
1. Le restore s'appuie uniquement sur des coordonnées absolues de bureau virtuel.
2. Le scénario ne mémorisait pas la topologie écrans ni le moniteur d'origine de chaque fenêtre.
3. `ensure_rect_on_screen` ne corrige que l'off-screen, pas un mauvais écran "encore visible".
- **Actions menées**:
1. Ajout de métadonnées d'affichage dans les sauvegardes de scénario.
2. Détection d'un changement de topologie écrans au chargement.
3. Adaptation automatique du `rect` de chaque fenêtre vers le moniteur courant correspondant.
4. Warning explicite pour les anciens scénarios sans métadonnées écran.
5. Ajout d'un rappel de fin de restauration pour expliquer qu'un placement imparfait peut être normal tant qu'une resauvegarde n'a pas été faite après changement d'écrans.
- **Décision**: privilégier l'adaptation automatique déterministe, avec resauvegarde recommandée pour les scénarios legacy.
- **Statut**: `monitoring`

### DBG-20260324-03
- **Titre**: Terminal de logs sans code couleur et avec doublons visuels
- **Contexte**: lecture des logs dans l'UI Tk pendant sauvegarde/restauration.
- **Symptômes**: tous les messages apparaissent avec la même couleur et certains warnings sont visibles en double à l'écran.
- **Étapes de reproduction**:
1. Lancer une restauration produisant plusieurs niveaux de logs.
2. Observer le terminal intégré.
- **Résultat attendu**: chaque niveau de log est facile à distinguer visuellement et une même ligne consécutive n'est pas affichée deux fois.
- **Résultat observé**: terminal monochrome vert et duplication ponctuelle de certaines lignes de warning.
- **Hypothèses**:
1. Le widget `Text` n'utilise aucun tag de couleur par niveau.
2. La redirection stdout/stderr insère les fragments bruts sans filtrage de doublons consécutifs.
- **Actions menées**:
1. Ajout de tags couleur par niveau dans `window_manager.pyw`.
2. Classification des fragments de log lors de l'insertion dans le terminal.
3. Suppression des doublons consécutifs exacts au niveau de l'affichage UI.
- **Décision**: améliorer la lisibilité sans modifier le format de log produit par `Logger`.
- **Statut**: `monitoring`

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
