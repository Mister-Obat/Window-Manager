import os
from .settings import SettingsManager
from .storage import LayoutStorage
from .scanner import WindowScanner
from .matcher import WindowMatcher
from .restorer import WindowRestorer
from .automation import is_incognito
from .logger import Logger
from .utils import capture_display_profile, get_rect_display_context

class WindowManagerEngine:
    def __init__(self, layout_file_path):
        self.layout_file = layout_file_path
        self.settings_file = os.path.join(os.path.dirname(self.layout_file), "settings.json")
        
        # Initialize Subsystems
        self.settings_manager = SettingsManager(self.settings_file)
        self.storage = LayoutStorage(self.layout_file)
        self.scanner = WindowScanner(self.settings_manager)
        self.matcher = WindowMatcher()
        self.restorer = WindowRestorer(self.settings_manager, self.scanner, self.matcher, self.storage)

    # --- Property Proxies for Backward Compatibility with UI ---
    @property
    def settings(self):
        return self.settings_manager.settings

    @settings.setter
    def settings(self, value):
        self.settings_manager.settings = value

    @property
    def layouts(self):
        return self.storage.layouts

    @layouts.setter
    def layouts(self, value):
        self.storage.layouts = value

    # --- Facade Methods ---
    def load_settings(self):
        return self.settings_manager.load_settings()

    def save_settings(self, new_settings):
        self.settings_manager.save_settings(new_settings)

    def load_layouts(self):
        return self.storage.load_layouts()

    def save_layout(self, scenario_name, overrides=None):
        self.scanner.clear_cache()
        Logger.title(f"Sauvegarde : {scenario_name}")
        
        # Determine Settings (Passed Overrides > Storage per-scenario > Global)
        # Usage: UI passes Slot Settings as overrides.
        final_settings = overrides
        if not final_settings:
             final_settings = self.storage.get_layout_settings(scenario_name)
        
        try:
            Logger.info("Scan des fenêtres ouvertes...")
            Logger.info("Début de l'analyse détaillée...")
            # Pass local settings as overrides
            windows = self.scanner.get_target_windows(detailed_scan=True, allow_peeking=True, overrides=final_settings)
            display_profile = capture_display_profile()
            
            Logger.info("Fin de l'analyse détaillée.")
            
            layout_data = []
            for w in windows:
                key = w["target_key"]
                if w["folder_path"]: key = "File Explorer"
                
                is_priv = w.get("is_incognito", False)
                title_log = w['title']
                if is_priv:
                    title_log = f"[PRIVÉ] {title_log}"
                
                Logger.info(f"Capturé : {title_log}")
                
                layout_data.append({
                    "title_pattern": key, 
                    "exact_title": w["title"], 
                    "rect": w["rect"],
                    "show_cmd": w["show_cmd"],
                    "cmdline": w["cmdline"],
                    "cwd": w["cwd"],
                    "url": w["url"],
                    "folder_path": w["folder_path"],
                    "is_incognito": is_priv,
                    "display": get_rect_display_context(w["rect"], display_profile),
                })

            self.storage.set_layout(scenario_name, {
                "windows": layout_data,
                "settings": final_settings or {},
                "display_profile": display_profile,
            })
            Logger.success(f"Sauvegarde terminée ({len(layout_data)} fenêtres)")
            return True
        except Exception as e:
            Logger.error(f"Erreur lors de la sauvegarde : {e}")
            return False

    def rename_layout(self, old_name, new_name):
        success = self.storage.rename_layout(old_name, new_name)
        if success:
            Logger.success(f"Renommé : '{old_name}' -> '{new_name}'")
        return success

    def restore_layout(self, scenario_name):
        return self.restorer.restore_layout(scenario_name)

    # --- Expose Helpers if needed ---
    def normalize_url(self, url):
        from .utils import normalize_url
        return normalize_url(url)

    def get_layout_settings(self, name):
        return self.storage.get_layout_settings(name)

    def set_layout_settings(self, name, settings):
        return self.storage.set_layout_settings(name, settings)

    def get_slot_settings(self, index):
        return self.settings_manager.get_slot_settings(index)

    def set_slot_settings(self, index, settings):
        return self.settings_manager.set_slot_settings(index, settings)

    def get_ui_slots(self):
        return self.settings_manager.get_ui_slots()
        
    def set_ui_slot(self, index, name):
        return self.settings_manager.set_ui_slot(index, name)
