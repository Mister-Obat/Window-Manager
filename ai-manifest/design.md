# Design Notes: Window Manager

## Principes UI
- Interface simple orientée actions: sauvegarder, restaurer, options.
- Feedback utilisateur immédiat via logs d'étapes (`Logger`).
- Le terminal de logs doit distinguer visuellement les niveaux (`info`, `warn`, `error`, `success`, `debug`) et éviter les doublons visibles inutiles.
- Thème sombre uniforme.

## Principes UX
- Priorité à la fiabilité perçue: chaque phase de restauration est visible.
- Tolérance aux apps lentes: waits/retries plutôt que fail rapide.
- Éviter les effets visuels indésirables pendant scan/restauration (peek minimal).

## Règles d'évolution
- Préférer patchs localisés dans `scanner`/`matcher`/`restorer` plutôt que refonte globale.
- Toute nouvelle heuristique doit être désactivable par contexte si risque de régression.
