import tkinter as tk
from tkinter import ttk, messagebox
import win32gui
import win32process
import os
import threading
from ctypes import windll, byref, sizeof, c_int
import winreg
import sys
import webbrowser
from wm_engine.engine import WindowManagerEngine

# Configuration
LAYOUT_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "layouts.json")

class WindowLayoutManagerUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Window Manager") 
        self.root.geometry("540x600")
        self.root.resizable(False, False)
        
        # Engine Initialization
        self.engine = WindowManagerEngine(LAYOUT_FILE)
        
        # Center Window
        self.center_window()
        self.apply_dark_title_bar()
        
        # Redirect stdout
        self.old_stdout = sys.stdout

        # Set App Icon
        base_path = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(base_path, "logo.ico")
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(default=icon_path)
                
                # FORCE ICON via Win32 API (WM_SETICON)
                # This fixes the Taskbar icon when AppID doesn't suffice
                import win32con
                
                hwnd = windll.user32.GetParent(self.root.winfo_id())
                
                # Load Small and Big icons
                hicon_big = win32gui.LoadImage(0, icon_path, win32con.IMAGE_ICON, 48, 48, win32con.LR_LOADFROMFILE)
                hicon_small = win32gui.LoadImage(0, icon_path, win32con.IMAGE_ICON, 16, 16, win32con.LR_LOADFROMFILE)
                
                if hicon_big:
                    win32gui.SendMessage(hwnd, win32con.WM_SETICON, win32con.ICON_BIG, hicon_big)
                if hicon_small:
                    win32gui.SendMessage(hwnd, win32con.WM_SETICON, win32con.ICON_SMALL, hicon_small)
                    
            except Exception as e:
                print(f"Icon Error: {e}") # Non-blocking
        else:
            messagebox.showwarning("Icone Manquante", f"Logo introuvable : {icon_path}")

        self.entries = [] 
        self.original_names = [] 

        # Modern UI Colors
        self.colors = {
            "bg": "#121212",           # Very dark grey/black background
            "surface": "#1e1e1e",      # Card/Container background
            "fg": "#e0e0e0",           # Soft white text
            "fg_sub": "#a0a0a0",       # Subtitle text
            "accent": "#4cc2ff",       # Modern blue accent
            "btn_bg": "#2d2d2d",       # Button normal
            "btn_hover": "#3d3d3d",    # Button hover
            "btn_primary": "#0078d4",  # Primary action
            "btn_primary_hover": "#0086e0"
        }

        self.create_widgets()
        
        # Overlay for status
        self.create_overlay()

    def center_window(self):
        self.root.update_idletasks()
        self.default_width = 500
        self.log_height = 180
        self.default_height = 460 + self.log_height
        x = (self.root.winfo_screenwidth() // 2) - (self.default_width // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.default_height // 2)
        self.root.geometry(f'{self.default_width}x{self.default_height}+{x}+{y}')

    def apply_dark_title_bar(self, window=None):
        target = window if window else self.root
        target.update_idletasks()
        try:
            window_handle = windll.user32.GetParent(target.winfo_id())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19
            if windll.dwmapi.DwmSetWindowAttribute(window_handle, DWMWA_USE_IMMERSIVE_DARK_MODE, byref(c_int(1)), sizeof(c_int(1))) != 0:
                windll.dwmapi.DwmSetWindowAttribute(window_handle, DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1, byref(c_int(1)), sizeof(c_int(1)))
        except Exception:
            pass

    def is_startup_enabled(self):
        # 1. Check VBS (Priority)
        startup_folder = os.path.join(os.getenv('APPDATA'), r'Microsoft\Windows\Start Menu\Programs\Startup')
        vbs_path = os.path.join(startup_folder, "WindowManager.vbs")
        if os.path.exists(vbs_path):
            return True

        # 2. Check Legacy Registry
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
            winreg.QueryValueEx(key, "WindowManager")
            winreg.CloseKey(key)
            return True
        except WindowsError:
            return False

    # toggle_startup, update_switch_ui, animate_switch removed (moved logic to save_and_close)

    def create_widgets(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure Colors & Fonts
        bg_color = self.colors["bg"]
        surface_color = self.colors["surface"]
        fg_color = self.colors["fg"]
        accent_color = self.colors["accent"]
        
        self.root.configure(bg=bg_color)
        
        # General Styles
        style.configure("TFrame", background=bg_color)
        style.configure("Card.TFrame", background=surface_color)
        
        style.configure("TLabel", background=bg_color, foreground=fg_color, font=("Segoe UI", 10))
        style.configure("Card.TLabel", background=surface_color, foreground=fg_color, font=("Segoe UI", 10))
        
        style.configure("Header.TLabel", font=("Segoe UI Variable Display", 20, "bold"), foreground=accent_color, background=bg_color)
        style.configure("Sub.TLabel", font=("Segoe UI", 10), foreground=self.colors["fg_sub"], background=bg_color)
        style.configure("Note.TLabel", font=("Segoe UI", 9), foreground=self.colors["fg_sub"], background=bg_color)

        # Entry Style
        style.configure("TEntry", fieldbackground=self.colors["surface"], foreground=fg_color, 
                        bordercolor=self.colors["surface"], lightcolor=self.colors["surface"], darkcolor=self.colors["surface"],
                        borderwidth=0, padding=10)
        
        # Button Styles
        # Save (Secondary)
        style.configure("Secondary.TButton", font=("Segoe UI", 9), background=self.colors["btn_bg"], foreground=fg_color, borderwidth=0, padding=6)
        style.map("Secondary.TButton", 
                  background=[("active", self.colors["btn_hover"]), ("pressed", self.colors["surface"])],
                  foreground=[("active", "white")])

        # Footer Checkbutton
        style.configure("Footer.TCheckbutton", background=bg_color, foreground=self.colors["fg_sub"], font=("Segoe UI", 9))
        style.map("Footer.TCheckbutton",
            background=[("active", bg_color)],
            foreground=[("active", self.colors["fg"])]
        )
        
        # Load (Primary)
        style.configure("Primary.TButton", font=("Segoe UI", 9, "bold"), background=self.colors["btn_primary"], foreground="white", borderwidth=0, padding=6)
        style.map("Primary.TButton", 
                  background=[("active", self.colors["btn_primary_hover"])])

        # Main Layout Container
        main_container = ttk.Frame(self.root, style="TFrame")
        main_container.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)

# Footer switch removed

        # Footer Frame
        footer_frame = ttk.Frame(main_container, style="TFrame")
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(20, 0))

        # Donate Button (Left)
        donate_btn = ttk.Button(footer_frame, text="Faire un don ❤️", style="Secondary.TButton",
                                command=self.open_donation, cursor="hand2")
        donate_btn.pack(side=tk.LEFT)

        # Startup Toggle (Right)
        self.startup_var = tk.BooleanVar(value=self.is_startup_enabled())
        startup_chk = ttk.Checkbutton(footer_frame, text="Lancer au démarrage", variable=self.startup_var, 
                                      style="Footer.TCheckbutton", command=self.toggle_startup_direct)
        startup_chk.pack(side=tk.RIGHT)
        
        # Actions Container (Header + Scenarios) - This will be covered by the overlay
        self.actions_container = ttk.Frame(main_container, style="TFrame")
        self.actions_container.pack(fill=tk.BOTH, expand=True)

        # Header Section
        header_frame = ttk.Frame(self.actions_container, style="TFrame")
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        title = ttk.Label(header_frame, text="Scénarios", style="Header.TLabel")
        title.pack(anchor="w")
        
        subtitle = ttk.Label(header_frame, text="Gérez votre espace de travail efficacement", style="Sub.TLabel")
        subtitle.pack(anchor="w", pady=(2, 0))

        # Scenarios List
        self.slots_frame = ttk.Frame(self.actions_container, style="TFrame")
        self.slots_frame.pack(fill=tk.BOTH, expand=True)

        # Log Area (Always visible)
        self.log_container = ttk.Frame(main_container, style="TFrame")
        self.log_container.pack(fill=tk.BOTH, expand=True, pady=(15, 0))
        
        self.log_text = tk.Text(self.log_container, bg="#000000", fg="#00ff00", 
                                font=("Consolas", 9), wrap=tk.WORD, state=tk.DISABLED,
                                bd=0, highlightthickness=1, highlightbackground="#333333", height=10)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._current_log_line = ""
        self._current_log_line_start = None
        self._last_completed_log_line = None
        self._configure_log_tags()
        
        scrollbar = tk.Scrollbar(self.log_container, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        sys.stdout = RedirectText(self) 
        sys.stderr = RedirectText(self)

        print(">>> Système de logs initialisé.")
        print(">>> Prêt.")

        # START MIGRATION / LOAD LOGIC
        ui_slots = self.engine.get_ui_slots()
        
        # If no slots defined but we have layouts (Legacy Migration)
        if not ui_slots and self.engine.layouts:
             # Sort keys to have deterministic migration
             existing_keys = sorted(list(self.engine.layouts.keys()))
             ui_slots = []
             for k in existing_keys:
                 ui_slots.append(k)
             # Pad remaining
             while len(ui_slots) < 5: ui_slots.append("")
             
             # Save immediate migration
             for idx, name in enumerate(ui_slots):
                 self.engine.set_ui_slot(idx, name)
        
        # Determine effective list
        effective_slots = ui_slots if ui_slots else [""] * 5
        while len(effective_slots) < 5: effective_slots.append("")

        for i in range(5):
            val = effective_slots[i]
            default_text = val if val else f"Scénario {i+1}"
            
            # Check if this name actually points to a real layout
            # If standard "Scénario X" and not in layouts, it's just placeholder.
            pass
                
            # Card Row
            row_frame = ttk.Frame(self.slots_frame, style="Card.TFrame") 
            row_frame.pack(fill=tk.X, pady=6, ipady=3)
            
            # Decoration
            strip = tk.Frame(row_frame, bg=accent_color, width=4)
            strip.pack(side=tk.LEFT, fill=tk.Y)
            
            # Input Area (Custom Styling)
            content_frame = ttk.Frame(row_frame, style="Card.TFrame", padding=(10, 5, 5, 5))
            content_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # Custom Entry Border Container
            entry_container = tk.Frame(content_frame, bg="#3d3d3d", padx=1, pady=1) # Border color
            entry_container.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=1)
            
            entry_inner = tk.Frame(entry_container, bg="#252526") # Inner bg
            entry_inner.pack(fill=tk.BOTH, expand=True)

            entry = tk.Entry(entry_inner, bg="#252526", fg=fg_color, 
                             insertbackground="white", bd=0, font=("Segoe UI", 10))
            entry.insert(0, default_text)
            entry.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)
            
            entry.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)
            
            self.entries.append(entry)
            # Store original name for Rename detection
            # If it's a generic "Scenario X" that doesn't exist in layouts, we treat it as empty/new.
            is_saved = default_text in self.engine.layouts
            self.original_names.append(default_text if is_saved else None)
            
            # Action Buttons
            action_frame = ttk.Frame(row_frame, style="Card.TFrame", padding=(0, 0, 10, 0))
            action_frame.pack(side=tk.RIGHT, fill=tk.Y)
            
            save_btn = ttk.Button(action_frame, text="SAUVER", style="Secondary.TButton", 
                                  command=lambda idx=i: self.save_from_ui(idx), cursor="hand2")
            save_btn.pack(side=tk.LEFT, padx=(0, 8))
            
            load_btn = ttk.Button(action_frame, text="CHARGER", style="Primary.TButton", 
                                  command=lambda idx=i: self.load_from_ui(idx), cursor="hand2")
            load_btn.pack(side=tk.LEFT, padx=(0, 8))

            # Options Button (Right of Charger)
            opt_btn = ttk.Button(action_frame, text="⚙️", style="Secondary.TButton", width=3,
                                  command=lambda idx=i: self.open_scenario_options(idx), cursor="hand2")
            opt_btn.pack(side=tk.LEFT)

    def create_overlay(self):
        # Overlay covering ONLY the actions container
        self.overlay_frame = tk.Frame(self.actions_container, bg="#1a1a1a") 
        self.overlay_frame.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.overlay_frame.place_forget() 

        container = tk.Frame(self.overlay_frame, bg="#1a1a1a")
        container.place(relx=0.5, rely=0.5, anchor="center")

        self.overlay_label = tk.Label(container, text="Traitement en cours...", 
                                      font=("Segoe UI", 16, "bold"), bg="#1a1a1a", fg=self.colors["accent"])
        self.overlay_label.pack(pady=(0, 10))
        
        self.overlay_sub = tk.Label(container, text="Veuillez patienter", 
                                    font=("Segoe UI", 11), bg="#1a1a1a", fg="#cccccc")
        self.overlay_sub.pack()

    def show_overlay(self, message):
        self.overlay_label.config(text=message)
        self.overlay_frame.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.overlay_frame.lift()
        self.root.update()

    def hide_overlay(self):
        self.overlay_frame.place_forget()
        self.root.update()

    def open_donation(self):
        webbrowser.open("https://www.paypal.com/paypalme/creaprisme")

    def toggle_startup_direct(self):
        should_run = self.startup_var.get()
        startup_folder = os.path.join(os.getenv('APPDATA'), r'Microsoft\Windows\Start Menu\Programs\Startup')
        vbs_path = os.path.join(startup_folder, "WindowManager.vbs")

        if should_run:
            try:
                exe = sys.executable.replace("python.exe", "pythonw.exe")
                script = os.path.abspath(__file__)
                working_dir = os.path.dirname(script)
                
                # Escape quotes for VBS string
                cmd_vbs = f'"{exe}" "{script}"'.replace('"', '""')
                
                vbs_content = (
                    'Set WshShell = CreateObject("WScript.Shell")\n'
                    f'WshShell.CurrentDirectory = "{working_dir}"\n'
                    f'WshShell.Run "{cmd_vbs}", 0, False\n'
                )
                
                # Use utf-16 to handle special characters (e.g., 'Créations') correctly in VBS
                with open(vbs_path, "w", encoding="utf-16") as f:
                    f.write(vbs_content)
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible d'activer le démarrage : {e}")
                self.startup_var.set(False) # Revert UI
        else:
            if os.path.exists(vbs_path):
                try:
                    os.remove(vbs_path)
                except Exception as e:
                    messagebox.showerror("Erreur", f"Impossible de désactiver le démarrage : {e}")
                    self.startup_var.set(True) # Revert UI


    def open_scenario_options(self, index):
        name = self.entries[index].get().strip()
        # Name is not strictly required for Slot settings but good for label
        display_name = name if name else f"Scénario {index+1}"

        # Prepare context (Slot Based)
        defaults = self.engine.settings
        overrides = self.engine.get_slot_settings(index)
        
        # Merge to get effective initial state
        def val(key, default=None):
            if key in overrides: return overrides[key]
            return defaults.get(key, default)

        # Create Dialog
        dlg = tk.Toplevel(self.root)
        dlg.title(f"Options : {display_name}")
        dlg.geometry("400x500") 
        dlg.config(bg=self.colors["bg"])
        dlg.resizable(False, False)
        
        dlg.after(10, lambda: self.apply_dark_title_bar(dlg))
        
        # Center
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 200
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 250
        dlg.geometry(f"400x500+{x}+{y}")
        
        dlg.transient(self.root)
        dlg.grab_set()

        # Vars
        vars = {
            "precise_urls": tk.BooleanVar(value=val("precise_urls", True)),
            "restore_minimized": tk.BooleanVar(value=val("restore_minimized", True)),
            "ignore_folders": tk.BooleanVar(value=val("ignore_folders", False)),
            "ignore_chrome": tk.BooleanVar(value=val("ignore_chrome", False)),
            "ignore_firefox": tk.BooleanVar(value=val("ignore_firefox", False)),
            "ignore_others": tk.BooleanVar(value=val("ignore_others", False)),
        }

        # Styles
        style = ttk.Style()
        style.configure("Opt.TCheckbutton", background=self.colors["bg"], foreground=self.colors["fg"], font=("Segoe UI", 10))
        style.map("Opt.TCheckbutton",
            background=[("active", self.colors["bg"]), ("disabled", self.colors["bg"])],
            foreground=[("active", self.colors["fg"]), ("disabled", self.colors["fg_sub"])]
        )
        style.configure("Title.TLabel", background=self.colors["bg"], foreground=self.colors["accent"], font=("Segoe UI Variable Display", 18, "bold"))
        style.configure("Section.TLabel", background=self.colors["bg"], foreground=self.colors["accent"], font=("Segoe UI", 11, "bold"))

        main = ttk.Frame(dlg, style="TFrame", padding=30)
        main.pack(fill=tk.BOTH, expand=True)

        header = ttk.Label(main, text=f"Réglages Ligne {index+1}", style="Title.TLabel")
        header.pack(anchor="w", pady=(0, 5))
        
        sub = ttk.Label(main, text=display_name, style="Sub.TLabel", font=("Segoe UI", 12))
        sub.pack(anchor="w", pady=(0, 20))

        # --- Comportement ---
        ttk.Label(main, text="COMPORTEMENT", style="Section.TLabel").pack(anchor="w", pady=(0, 10))
        
        ttk.Checkbutton(main, text="Précision des URLs (Browsers)", variable=vars["precise_urls"], style="Opt.TCheckbutton").pack(anchor="w")
        ttk.Checkbutton(main, text="Restaurer les fenêtres réduites", variable=vars["restore_minimized"], style="Opt.TCheckbutton").pack(anchor="w", pady=(5, 0))

        ttk.Separator(main, orient='horizontal').pack(fill='x', pady=20)

        # --- Filtres ---
        ttk.Label(main, text="FILTRES (IGNORER)", style="Section.TLabel").pack(anchor="w", pady=(0, 10))
        
        grid_frame = tk.Frame(main, bg=self.colors["bg"])
        grid_frame.pack(fill=tk.X)
        
        def make_chk(parent, text, var, row, col):
            f = tk.Frame(parent, bg=self.colors["bg"])
            f.grid(row=row, column=col, sticky="w", pady=8, padx=(0, 10))
            ttk.Checkbutton(f, text=text, variable=var, style="Opt.TCheckbutton").pack()

        make_chk(grid_frame, "Dossiers", vars["ignore_folders"], 0, 0)
        make_chk(grid_frame, "Google Chrome", vars["ignore_chrome"], 0, 1)
        make_chk(grid_frame, "Mozilla Firefox", vars["ignore_firefox"], 1, 0)
        make_chk(grid_frame, "Autres programmes", vars["ignore_others"], 1, 1)

        grid_frame.columnconfigure(0, weight=1)
        grid_frame.columnconfigure(1, weight=1)

        # Save Action
        footer = tk.Frame(main, bg=self.colors["bg"])
        footer.pack(fill=tk.X, side=tk.BOTTOM)

        def save_local():
            # Export values to SLOT settings
            final_settings = {k: v.get() for k, v in vars.items()}
            self.engine.set_slot_settings(index, final_settings)
            dlg.destroy()
            messagebox.showinfo("Succès", f"Paramètres sauvegardés pour la ligne {index+1}.", parent=self.root)

        ttk.Button(footer, text="ENREGISTRER", style="Primary.TButton", command=save_local, cursor="hand2").pack(side=tk.RIGHT, ipadx=10)

    def save_from_ui(self, index):
        new_name = self.entries[index].get().strip()
        if not new_name:
             messagebox.showwarning("Nom manquant", "Veuillez entrer un nom pour ce scénario.", parent=self.root)
             return

        old_name = self.original_names[index]
        slot_settings = self.engine.get_slot_settings(index)

        # Case: Renaming an EXISTING layout
        if old_name and new_name != old_name and old_name in self.engine.layouts:
            # Check if target name also exists
            if new_name in self.engine.layouts:
                 pass 
            
            # Show Choice Dialog
            choice = self.ask_rename_or_save(old_name, new_name)
            if choice == "cancel": return
            
            if choice == "rename":
                if new_name in self.engine.layouts:
                     if not messagebox.askyesno("Attention", f"Le nom '{new_name}' est déjà pris.\nÉcraser ?", parent=self.root):
                         return
                
                # RENAME OPERATION
                success = self.engine.rename_layout(old_name, new_name)
                if success:
                    self.original_names[index] = new_name # Update memory
                    self.engine.set_ui_slot(index, new_name) # Lock slot
                return

        # Case: Standard Save (New or Overwrite)
        if new_name in self.engine.layouts:
            should_ask = True
            if old_name == new_name:
                 pass
            
            if should_ask:
                if not messagebox.askyesno("Confirmation", f"Le scénario '{new_name}' existe déjà.\nVoulez-vous l'écraser ?", parent=self.root):
                    return

        self.save_layout_thread(new_name, overrides=slot_settings)
        self.original_names[index] = new_name
        self.engine.set_ui_slot(index, new_name) # Lock slot


    def ask_rename_or_save(self, old_name, new_name):
        """ Custom Dialog for Rename vs Save As """
        diag = tk.Toplevel(self.root)
        diag.title("Action requise")
        diag.config(bg=self.colors["bg"])
        diag.resizable(False, False)
        
        self.apply_dark_title_bar(diag)
        
        # Center
        w, h = 400, 220
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (w // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (h // 2)
        diag.geometry(f"{w}x{h}+{x}+{y}")
        diag.transient(self.root)
        diag.grab_set()
        
        diag.result = "cancel"

        lbl = ttk.Label(diag, text=f"Vous avez renommé '{old_name}' en '{new_name}'.", 
                        style="Card.TLabel", font=("Segoe UI", 11), justify="center")
        lbl.pack(pady=(25, 5))
        
        lbl2 = ttk.Label(diag, text="Que voulez-vous faire ?", style="Card.TLabel", font=("Segoe UI", 10, "bold"))
        lbl2.pack(pady=(0, 20))

        def set_rename():
            diag.result = "rename"
            diag.destroy()

        def set_save():
            diag.result = "save"
            diag.destroy()

        btn_rename = ttk.Button(diag, text="RENOMMER Seulement\n(Garder les fenêtres sauvegardées)", 
                                style="Secondary.TButton", command=set_rename)
        btn_rename.pack(fill=tk.X, padx=30, pady=5)
        
        btn_save = ttk.Button(diag, text="NOUVELLE SAUVEGARDE\n(Capturer les fenêtres actuelles)", 
                              style="Primary.TButton", command=set_save)
        btn_save.pack(fill=tk.X, padx=30, pady=5)

        diag.wait_window()
        return diag.result

    def load_from_ui(self, index):
        name = self.entries[index].get().strip()
        if name:
            self.restore_layout_thread(name)
        else:
             messagebox.showwarning("Nom manquant", "Veuillez entrer un nom pour ce scénario.")
    
    def save_layout_thread(self, scenario_name, overrides=None):
        self.show_overlay(f"Sauvegarde de\n'{scenario_name}'")
        t = threading.Thread(target=self._save_layout_wrapped, args=(scenario_name, overrides))
        t.start()
        self.check_thread(t)

    def restore_layout_thread(self, scenario_name):
        self.show_overlay(f"Restauration de\n'{scenario_name}'")
        t = threading.Thread(target=self._restore_layout_wrapped, args=(scenario_name,))
        t.start()
        self.check_thread(t)

    def check_thread(self, thread):
        if thread.is_alive():
            self.root.after(100, lambda: self.check_thread(thread))
        else:
            self.hide_overlay()

    def _save_layout_wrapped(self, scenario_name, overrides):
        success = self.engine.save_layout(scenario_name, overrides)
        if not success:
            # We could use message box mainly if it wasn't threaded this way, 
            # but since we are in a thread we'd need to schedule it back to main loop.
            # Ideally extend show_overlay to show "Success" or "Error" briefly for better UX.
            pass

    def _restore_layout_wrapped(self, scenario_name):
        self.engine.restore_layout(scenario_name)

    def write_to_log(self, message):
        # Thread-safe log update using after()
        self.root.after(0, self._append_log_text, message)

    def _configure_log_tags(self):
        self.log_text.configure(fg="#d0d0d0")
        self.log_text.tag_configure("log_default", foreground="#d0d0d0")
        self.log_text.tag_configure("log_info", foreground="#8fd3ff")
        self.log_text.tag_configure("log_success", foreground="#7fe39d")
        self.log_text.tag_configure("log_warn", foreground="#ffd166")
        self.log_text.tag_configure("log_error", foreground="#ff7b72")
        self.log_text.tag_configure("log_title", foreground="#4cc2ff", font=("Consolas", 9, "bold"))
        self.log_text.tag_configure("log_debug", foreground="#b39ddb")
        self.log_text.tag_configure("log_private", foreground="#ffb86c")

    def _resolve_log_tag(self, segment, current_line):
        text = f"{current_line}{segment}".strip()
        if not text:
            return "log_default"

        if "[ERREUR]" in text or "[ECHEC]" in text:
            return "log_error"
        if "[!]" in text:
            return "log_warn"
        if "[DEBUG]" in text:
            return "log_debug"
        if text.startswith("---"):
            return "log_title"
        if "[OK]" in text:
            return "log_success"
        if "[PRIVÉ]" in text:
            return "log_private"
        if text.startswith(">") or text.startswith(">>>") or text.startswith("  >"):
            return "log_info"
        return "log_default"

    def _append_log_text(self, message):
        if not message:
            return

        normalized = message.replace("\r\n", "\n").replace("\r", "\n")
        self.log_text.config(state=tk.NORMAL)
        try:
            for segment in normalized.splitlines(keepends=True):
                if self._current_log_line_start is None:
                    self._current_log_line_start = self.log_text.index(tk.END)

                tag = self._resolve_log_tag(segment, self._current_log_line)
                self.log_text.insert(tk.END, segment, tag)
                self._current_log_line += segment

                if segment.endswith("\n"):
                    completed_line = self._current_log_line
                    if completed_line.strip() and completed_line == self._last_completed_log_line:
                        self.log_text.delete(self._current_log_line_start, tk.END)
                    elif completed_line.strip():
                        self._last_completed_log_line = completed_line

                    self._current_log_line = ""
                    self._current_log_line_start = None

            self.log_text.see(tk.END)
        finally:
            self.log_text.config(state=tk.DISABLED)
            self.log_text.update_idletasks() # Force UI refresh

# Update RedirectText to use the method
class RedirectText:
    def __init__(self, app):
        self.app = app

    def write(self, string):
        self.app.write_to_log(string)

    def flush(self):
        pass

if __name__ == "__main__":
    import ctypes
    try:
        # Explicitly set the AppUserModelID to a static string to ensure taskbar consistency
        myappid = 'obat.window_manager.v1' 
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception as e:
        print(f"Could not set AppUserModelID: {e}")

    try:
        # Try to set Per-Monitor V2 Awareness (Windows 10 1703+)
        # MPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
        ctypes.windll.user32.SetProcessDpiAwarenessContext(-4)
    except Exception:
        try:
            # Fallback to Per-Monitor Awareness (Windows 8.1+)
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                # Fallback to System Awareness (Windows Vista+)
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

    root = tk.Tk()
    app = WindowLayoutManagerUI(root)
    root.mainloop()
