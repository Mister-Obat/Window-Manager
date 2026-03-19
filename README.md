# Window Manager

**Window Manager** est une application desktop Windows qui sauvegarde puis restaure votre espace de travail.  
Elle mémorise la position, la taille, l'état et le contexte des fenêtres pour une restauration fiable après redémarrage ou relance d'applications. 

![Window Manager Main](screenshot.png)
![Window Manager Options](screenshot2.png)

## Fonctionnalités

*   **Sauvegarde/Restauration de scénarios** : géométrie (`rect`) + état (`normal`, `minimisé`, `maximisé`).
*   **Restauration intelligente** : matching des fenêtres déjà ouvertes puis lancement séquentiel des manquantes.
*   **Navigateurs pris en charge** : Chrome, Firefox et Edge (avec gestion du mode privé).
*   **Explorateur Windows** : restauration des dossiers ouverts.
*   **Réglages par ligne/scénario** : `restore_minimized`, `precise_urls`, filtres d'exclusion.
*   **Démarrage automatique** : option d'exécution au démarrage de Windows.

## Installation

### Prérequis
*   Windows 10 ou 11
*   Python 3.11+ recommandé

### Configuration
1.  Clonez le dépôt :
    ```bash
    git clone https://github.com/Mister-Obat/Window-Manager-2.git
    cd "Window Manager"
    ```

2.  Installez les dépendances :
    ```bash
    pip install -r requirements.txt
    ```

## Utilisation

Lancez l'application via le script principal :
```bash
py window_manager.pyw
```

Option alternative :
```bash
start_manager.bat
```

### Gestion des Scénarios
1.  **Sauvegarder** : Configurez votre espace de travail, nommez le scénario et cliquez sur **SAUVER**.
2.  **Restaurer** : Cliquez sur **CHARGER** pour rouvrir et repositionner vos applications.
3.  **Options** : Utilisez le bouton `⚙️` pour régler les overrides de la ligne/scénario.

## Technologies

*   **Python** / **Tkinter** (UI desktop)
*   **pywin32** (`win32gui`, `win32con`, `win32process`) pour l'intégration Win32
*   **psutil** pour les métadonnées process (`cmdline`, `cwd`, `exe`)
*   **uiautomation** pour l'inspection navigateur (URL / mode privé)

## License
Ce projet est distribué sous licence AGPL-3.0.

---
*Codé 100% par des IA, supervisé à l'arrache par Obat 😏*
