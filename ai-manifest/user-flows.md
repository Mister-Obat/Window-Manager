# User Flows: Window Manager

## Flux 1: Sauvegarder un scénario
1. L'utilisateur ouvre Window Manager.
2. Il choisit un slot/scénario.
3. Il clique `SAUVER`.
4. L'app scanne les fenêtres, filtre, puis écrit `layouts.json`.

## Flux 2: Restaurer un scénario
1. L'utilisateur clique `CHARGER` sur un slot.
2. Phase 1: scan + matching des fenêtres déjà ouvertes.
3. Phase 2: lancement des fenêtres manquantes.
4. Chaque fenêtre est repositionnée selon `rect` et `show_cmd`.

## Flux 3: Démarrage automatique
1. Option "Lancer au démarrage" activée.
2. Le script est exécuté via startup Windows.
3. L'utilisateur restaure son scénario de travail.

## Cas de vigilance
- Applications Electron/Chromium: éviter les forçages de repaint agressifs.
- Fenêtres minimisées: matching plus permissif mais contrôlé.
