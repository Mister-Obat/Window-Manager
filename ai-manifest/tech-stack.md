# Tech Stack Justification: Window Manager

## Runtime
- Python 3.11 (`pythonw` en usage desktop).
- UI via `tkinter` / `ttk`.

## Intégrations système
- `pywin32` (`win32gui`, `win32con`, `win32process`) pour enum/placement des fenêtres.
- `ctypes` + DWM API pour états cloaked / titlebar.
- `psutil` pour métadonnées process (`cmdline`, `cwd`, exe).

## Moteur applicatif
- Architecture modulaire `wm_engine/`:
  - `scanner`: découverte et collecte de fenêtres.
  - `matcher`: scoring/matching entre sauvegardé et courant.
  - `restorer`: lancement + placement + orchestration.
  - `storage` / `settings`: persistance JSON.
  - `automation`: extraction URL navigateur / path Explorer.

## Persistance
- `layouts.json`: scénarios sauvegardés (`windows` + `settings`).
- `settings.json`: paramètres globaux et slots UI.

## Contraintes techniques
- API Win32 non transactionnelle: nécessiter retries et timings courts.
- Certaines apps (Electron/Chromium) ont un cycle de rendu fragile après `SW_RESTORE` + repaint forcé.
