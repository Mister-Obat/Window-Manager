# 🪟 Window Layout Manager (Gestionnaire de Fenêtres)

Un utilitaire Windows élégant et puissant pour sauvegarder et restaurer instantanément la disposition de vos fenêtres. Idéal pour retrouver votre espace de travail parfait en un clic après un redémarrage.

![Window Manager Screenshot](https://via.placeholder.com/500x480?text=Window+Manager+UI) TODO: Ajouter une capture d'écran ici.

## ✨ Fonctionnalités

*   **⚡ Sauvegarde & Restauration Complète** : Mémorise la position, la taille et l'état de **toutes** vos fenêtres actives.
*   **📚 Support Universel** : Fonctionne avec n'importe quelle application Windows.
*   **🧠 Intelligence Contextuelle** :
    *   **Navigateurs Web** : Restaure les URLs spécifiques (Chrome, Firefox, Edge).
    *   **Explorateur de Fichiers** : Rouvre les dossiers exacts.
*   **🏗️ Gestion du Z-Order** : Restaure l'ordre de superposition des fenêtres (les fenêtres d'arrière-plan restent en arrière-plan).
*   **🚀 Lancement Automatique** : Option intégrée pour se lancer au démarrage de Windows (discret, sans console).
*   **🎨 Interface Moderne** : Thème sombre (Dark Mode), design compact et animations fluides.

## 🛠️ Prérequis

*   Windows 10 ou 11
*   Python 3.x

## 📦 Installation

1.  Clonez ce dépôt :
    ```bash
    git clone https://github.com/votre-username/window-manager.git
    cd window-manager
    ```

2.  Installez les dépendances :
    ```bash
    pip install -r requirements.txt
    ```

## 🚀 Utilisation

1.  Lancez l'application :
    ```bash
    py window_manager.pyw
    ```
    *(L'extension `.pyw` permet de lancer l'application sans fenêtre de console persistante)*.

2.  **Sauvegarder un scénario** :
    *   Disposez vos fenêtres comme vous le souhaitez.
    *   Entrez un nom pour votre scénario (ex: "Travail", "Gaming", "Streaming").
    *   Cliquez sur **SAUVER**.

3.  **Restaurer un scénario** :
    *   Cliquez sur **CHARGER** à côté du scénario désiré.
    *   L'application relancera les programmes manquants et repositionnera toutes les fenêtres.

4.  **Démarrage Auto** :
    *   Activez le switch "Lancer au démarrage" en bas de l'application pour qu'elle soit toujours prête.

## 🔧 Technologies

*   **Python** : Langage principal.
*   **Tkinter** : Interface graphique (GUI) native et légère.
*   **PyWin32 (win32gui, win32con)** : Interaction bas niveau avec l'API Windows (titres, positions, styles).
*   **UIAutomation** : Extraction avancée de données (URLs des navigateurs).
*   **WinReg** : Manipulation du registre pour le démarrage automatique.

## 📝 Auteur

Développé pour optimiser la productivité et la gestion du multitâche sur Windows.

## 📄 Licence

Open source sous licence **AGPL-3.0** pour usage personnel et non commercial.

Pour toute utilisation commerciale merci de me contacter.
📧 Mail : contact.creaprisme@gmail.com
