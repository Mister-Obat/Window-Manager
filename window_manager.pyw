import tkinter as tk
from tkinter import ttk, messagebox
import win32gui
import win32con
import win32process
import json
import os
import psutil
import subprocess
import time
import threading
import uiautomation as auto
import win32com.client
import urllib.parse
import ctypes
from ctypes import windll, byref, sizeof, c_int
import winreg
import sys

# Configuration
# Configuration
EXCLUDE_TITLES = [
    "Program Manager",
    "Microsoft Text Input Application",
    "Settings", # Windows Settings often behave oddly with simple restoration
]
# We no longer filter by TARGET_TITLES whitelist

LAYOUT_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "layouts.json")

class WindowLayoutManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Window Manager") 
        self.root.geometry("540x600")
        self.root.resizable(False, False)
        
        # Center Window
        self.center_window()
        self.apply_dark_title_bar()

        # Set App Icon
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.ico")
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(default=icon_path)
            except Exception:
                pass

        self.layouts = self.load_layouts()
        self.entries = [] 

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
        width = 500
        height = 480
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def apply_dark_title_bar(self):
        try:
            window_handle = windll.user32.GetParent(self.root.winfo_id())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19
            if windll.dwmapi.DwmSetWindowAttribute(window_handle, DWMWA_USE_IMMERSIVE_DARK_MODE, byref(c_int(1)), sizeof(c_int(1))) != 0:
                windll.dwmapi.DwmSetWindowAttribute(window_handle, DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1, byref(c_int(1)), sizeof(c_int(1)))
        except Exception:
            pass

    def is_startup_enabled(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
            winreg.QueryValueEx(key, "ObatWindowManager")
            winreg.CloseKey(key)
            return True
        except WindowsError:
            return False

    def toggle_startup(self):
        enabled = self.is_startup_enabled()
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
            if enabled:
                try:
                    winreg.DeleteValue(key, "ObatWindowManager")
                    self.switch_on = False
                except WindowsError:
                    pass
            else:
                exe = sys.executable.replace("python.exe", "pythonw.exe") # Use pythonw to hide console
                script = os.path.abspath(__file__)
                cmd = f'"{exe}" "{script}"'
                winreg.SetValueEx(key, "ObatWindowManager", 0, winreg.REG_SZ, cmd)
                self.switch_on = True
            winreg.CloseKey(key)
            self.update_switch_ui()
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de modifier le démarrage : {e}")

    def update_switch_ui(self):
        # We will animate the transition
        target_color = self.colors["accent"] if self.switch_on else "#3a3a3a"
        
        # New focused coordinates
        start_x = 14 if not self.switch_on else 32
        target_x = 32 if self.switch_on else 14
        
        current_coords = self.switch_canvas.coords(self.switch_circle)
        current_x = (current_coords[0] + current_coords[2]) / 2
        
        self.animate_switch(current_x, target_x, target_color)

    def animate_switch(self, current_x, target_x, target_color):
        step = 2 if target_x > current_x else -2
        
        # Clamp and finish
        if (step > 0 and current_x >= target_x) or (step < 0 and current_x <= target_x):
            self.switch_canvas.itemconfig(self.switch_rect, fill=target_color)
            self.switch_canvas.coords(self.switch_circle, target_x-7, 5, target_x+7, 19)
            return

        new_x = current_x + step
        self.switch_canvas.coords(self.switch_circle, new_x-7, 5, new_x+7, 19)
        self.switch_canvas.itemconfig(self.switch_rect, fill=target_color)
        
        self.root.after(10, lambda: self.animate_switch(new_x, target_x, target_color))

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
        
        # Load (Primary)
        style.configure("Primary.TButton", font=("Segoe UI", 9, "bold"), background=self.colors["btn_primary"], foreground="white", borderwidth=0, padding=6)
        style.map("Primary.TButton", 
                  background=[("active", self.colors["btn_primary_hover"])])

        # Main Layout Container
        main_container = ttk.Frame(self.root, style="TFrame")
        main_container.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)

        # Footer with Startup Switch
        footer_frame = ttk.Frame(main_container, style="TFrame")
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(20, 0))
        
        # Switch container
        switch_container = tk.Frame(footer_frame, bg=self.colors["bg"])
        switch_container.pack(anchor="w")
        
        # Label
        lbl = ttk.Label(switch_container, text="Lancer au démarrage", style="Note.TLabel")
        lbl.pack(side=tk.LEFT, padx=(0, 10))
        
        # Custom Switch Widget
        self.switch_on = self.is_startup_enabled()
        # Compact canvas
        self.switch_canvas = tk.Canvas(switch_container, width=46, height=24, bg=self.colors["bg"], highlightthickness=0)
        self.switch_canvas.pack(side=tk.LEFT)
        
        # Draw switch parts
        self.switch_canvas.delete("all")
        
        # Pill background: Tighter, cleaner
        bg_color = self.colors["accent"] if self.switch_on else "#3a3a3a"
        # Center Y=12. Width=20 (radius 10). Line from x=12 to x=34.
        self.switch_rect = self.switch_canvas.create_line(14, 12, 32, 12, width=20, capstyle=tk.ROUND, fill=bg_color)
        
        # Knob: Radius 7 (Diameter 14). White.
        x_circle = 32 if self.switch_on else 14
        self.switch_circle = self.switch_canvas.create_oval(x_circle-7, 5, x_circle+7, 19, fill="white", outline="")
        
        self.switch_canvas.bind("<Button-1>", lambda e: self.toggle_startup())
        self.switch_canvas.config(cursor="hand2")
        
        # Header Section
        header_frame = ttk.Frame(main_container, style="TFrame")
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        title = ttk.Label(header_frame, text="Scénarios", style="Header.TLabel")
        title.pack(anchor="w")
        
        subtitle = ttk.Label(header_frame, text="Gérez votre espace de travail efficacement", style="Sub.TLabel")
        subtitle.pack(anchor="w", pady=(2, 0))

        # Scenarios List
        self.slots_frame = ttk.Frame(main_container, style="TFrame")
        self.slots_frame.pack(fill=tk.BOTH, expand=True)

        existing_keys = sorted(list(self.layouts.keys()))
        
        for i in range(5):
            default_text = f"Scénario {i+1}"
            if i < len(existing_keys):
                default_text = existing_keys[i]
                
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
            
            self.entries.append(entry)
            
            # Action Buttons
            action_frame = ttk.Frame(row_frame, style="Card.TFrame", padding=(0, 0, 10, 0))
            action_frame.pack(side=tk.RIGHT, fill=tk.Y)
            
            save_btn = ttk.Button(action_frame, text="SAUVER", style="Secondary.TButton", 
                                  command=lambda idx=i: self.save_from_ui(idx), cursor="hand2")
            save_btn.pack(side=tk.LEFT, padx=(5, 8))
            
            load_btn = ttk.Button(action_frame, text="LANCER", style="Primary.TButton", 
                                  command=lambda idx=i: self.load_from_ui(idx), cursor="hand2")
            load_btn.pack(side=tk.LEFT)

    def create_overlay(self):
        # Overlay covering everything with high contrast
        self.overlay_frame = tk.Frame(self.root, bg="#1a1a1a") 
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

    def save_from_ui(self, index):
        name = self.entries[index].get().strip()
        if name:
            self.save_layout_thread(name)
        else:
             messagebox.showwarning("Nom manquant", "Veuillez entrer un nom pour ce scénario.")

    def load_from_ui(self, index):
        name = self.entries[index].get().strip()
        if name:
            self.restore_layout_thread(name)
        else:
             messagebox.showwarning("Nom manquant", "Veuillez entrer un nom pour ce scénario.")
    
    def save_layout_thread(self, scenario_name):
        self.show_overlay(f"Sauvegarde de\n'{scenario_name}'")
        t = threading.Thread(target=self._save_layout_wrapped, args=(scenario_name,))
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

    def _save_layout_wrapped(self, scenario_name):
        self.save_layout(scenario_name)

    def _restore_layout_wrapped(self, scenario_name):
        self.restore_layout(scenario_name)

    def load_layouts(self):
        if os.path.exists(LAYOUT_FILE):
            try:
                with open(LAYOUT_FILE, "r") as f:
                    return json.load(f)
            except:
                return {}
        return {}

    def save_layouts_to_disk(self):
        with open(LAYOUT_FILE, "w") as f:
            json.dump(self.layouts, f, indent=4)

    def extract_url_from_window(self, hwnd):
        try:
            window = auto.ControlFromHandle(hwnd)
            potential_names = [
                "Address and search bar", "Barre d'adresse et de recherche",
                "Address", "Adresse",
                "Search with Google or enter address",
                "Rechercher avec Google ou saisir une adresse"
            ]
            for name in potential_names:
                edit = window.EditControl(Name=name, searchDepth=15) 
                if edit.Exists(maxSearchSeconds=0.05):
                    val = edit.GetValuePattern().Value
                    if val: return val

            try:
                edit = window.EditControl(RegexName=".*(http|https|www|localhost|://).*", searchDepth=15)
                if edit.Exists(maxSearchSeconds=0.05):
                    val = edit.GetValuePattern().Value
                    if val: return val
            except:
                pass

            count = 0
            for control, depth in auto.WalkControl(window, maxDepth=12):
                if control.ControlTypeName == "EditControl":
                    try:
                        name = control.Name
                        if "http" in name or "www." in name or "localhost" in name:
                             if control.GetPattern(auto.PatternId.ValuePattern):
                                return control.GetValuePattern().Value
                        if control.GetPattern(auto.PatternId.ValuePattern):
                            val = control.GetValuePattern().Value
                            if val and ("." in val or "http" in val or "localhost" in val):
                                return val
                    except:
                        pass
                if count > 2000: break
                count += 1
        except Exception:
            pass
        return None

    def extract_path_from_explorer(self, hwnd):
        # Uses Shell.Application (COM) to reliably get the path from Explorer
        # regardless of UI language or Windows version (10/11)
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            for window in shell.Windows():
                try:
                    if window.HWND == hwnd:
                         # LocationURL is usually file:///C:/Path/To/Folder
                         url = window.LocationURL
                         if url.lower().startswith("file:///"):
                             # Decode URI
                             path_encoded = url[8:].replace("/", "\\")
                             path = urllib.parse.unquote(path_encoded)
                             if os.path.isdir(path):
                                 return path
                except:
                    pass
        except Exception:
            pass
        return None

    def normalize_url(self, url):
        if not url: return None
        if url.startswith("localhost") and not url.startswith("http"):
            return "http://" + url
        if not url.startswith("http") and not url.startswith("file") and "://" not in url:
            return "https://" + url
        return url

    def get_target_windows(self, detailed_scan=False):
        windows = []
        def enum_handler(hwnd, ctx):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if not title: return
                
                # Filter out system noise
                if title in EXCLUDE_TITLES: return
                
                class_name = win32gui.GetClassName(hwnd)
                
                # Basic Noise Filters
                if title == "Taskbar": return
                if title == "Mise en veille": return  # Example of potential system overlay

                is_explorer = (class_name == "CabinetWClass")
                
                # Check for "tool window" style which shouldn't be captured (usually)
                # style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
                # if style & win32con.WS_EX_TOOLWINDOW: return 
                # (Commented out for now as some apps use toolwindow specific ways, let's stick to title visibility first)

                # Capture logic
                try:
                    placement = win32gui.GetWindowPlacement(hwnd)
                    rec_placement = list(placement[4])
                    show_cmd = placement[1]
                except:
                    rec_placement = [0,0,0,0]
                    show_cmd = win32con.SW_SHOWNORMAL

                if show_cmd == win32con.SW_SHOWNORMAL:
                    try:
                        rect_actual = win32gui.GetWindowRect(hwnd)
                        rect = list(rect_actual)
                    except:
                        rect = rec_placement
                else:
                    rect = rec_placement
                
                # Filter out zero-size windows (ghost windows)
                w = rect[2] - rect[0]
                h = rect[3] - rect[1]
                if w < 10 or h < 10: return

                cmdline = []
                cwd = ""
                url = None
                folder_path = None
                
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    proc = psutil.Process(pid)
                    try:
                        cmdline = proc.cmdline()
                        cwd = proc.cwd()
                    except (psutil.AccessDenied, psutil.NoSuchProcess):
                        # Some system apps won't let us read cmdline. 
                        # We might not be able to relaunch them, but we can manage their window if open.
                        pass
                except:
                    pass
                
                if detailed_scan:
                    # Specific Heuristics for Data extraction
                    if "Chrome" in title or "Firefox" in title or "Edge" in title:
                            if not win32gui.IsIconic(hwnd):
                                url = self.extract_url_from_window(hwnd)
                    elif is_explorer and not win32gui.IsIconic(hwnd):
                        folder_path = self.extract_path_from_explorer(hwnd)

                windows.append({
                    "hwnd": hwnd, 
                    "title": title,
                    "target_key": "File Explorer" if is_explorer else title, # Strategy: Use Title as key by default
                    "rect": rect,
                    "show_cmd": show_cmd,
                    "cmdline": cmdline,
                    "cwd": cwd,
                    "url": url,
                    "folder_path": folder_path
                })
                return

        win32gui.EnumWindows(enum_handler, None)
        return windows

    def save_layout(self, scenario_name):
        windows = self.get_target_windows(detailed_scan=True)
        layout_data = []
        for w in windows:
            key = w["target_key"]
            if w["folder_path"]: key = "File Explorer"
            
            layout_data.append({
                "title_pattern": key, 
                "exact_title": w["title"], 
                "rect": w["rect"],
                "show_cmd": w["show_cmd"],
                "cmdline": w["cmdline"],
                "cwd": w["cwd"],
                "url": w["url"],
                "folder_path": w["folder_path"]
            })
        
        self.layouts[scenario_name] = layout_data
        self.save_layouts_to_disk()
        print(f"Saved {scenario_name}")

    def restore_layout(self, scenario_name):
        if scenario_name not in self.layouts:
            print(f"Not Found: {scenario_name}")
            return

        saved_windows = self.layouts[scenario_name]
        
        # 1. Launch Missing Apps (Bottom-to-Top)
        current_windows = self.get_target_windows(detailed_scan=True)
        initial_window_count = len(current_windows) # Snapshot current count
        
        current_inventory = {} 
        for w in current_windows:
            key = w["target_key"]
            if key not in current_inventory:
                current_inventory[key] = []
            current_inventory[key].append(w)
            
        simulated_inventory = {k: list(v) for k, v in current_inventory.items()}
        launched_count = 0

        # Processing saved_windows in REVERSE (Bottom -> Top)
        for saved in reversed(saved_windows):
            key = saved["title_pattern"]
            saved_url = saved.get("url")
            saved_folder = saved.get("folder_path")
            
            found = False
            if key in simulated_inventory:
                for i, w in enumerate(simulated_inventory[key]):
                    current_url = w.get("url")
                    current_folder = w.get("folder_path")
                    
                    if saved_folder and current_folder and saved_folder == current_folder:
                        simulated_inventory[key].pop(i)
                        found = True
                        break
                    elif saved_url and current_url and (saved_url in current_url or current_url in saved_url):
                         simulated_inventory[key].pop(i)
                         found = True
                         break
                    elif not saved_url and not saved_folder and not current_url and not current_folder:
                        simulated_inventory[key].pop(i)
                        found = True
                        break
                
                if not found:
                     for i, w in enumerate(simulated_inventory[key]):
                         if not saved_url and not saved_folder:
                             simulated_inventory[key].pop(i)
                             found = True
                             break
            
            if not found:
                print(f"Lancement de: {key}")
                cmdline = saved.get("cmdline")
                cwd = saved.get("cwd")
                url = saved.get("url")
                folder = saved.get("folder_path")
                
                try:
                    if folder:
                        subprocess.Popen(f'explorer.exe "{folder}"', shell=True)
                        launched_count += 1
                        time.sleep(1.0) # Folders can be slow too
                    elif ("Chrome" in key or "Firefox" in key or "Edge" in key):
                        browser_exe = None
                        if cmdline:
                            for arg in cmdline:
                                if arg.lower().endswith(".exe"):
                                    browser_exe = arg
                                    break
                        if not browser_exe and cmdline: browser_exe = cmdline[0]

                        if browser_exe:
                            if url: url = self.normalize_url(url)
                            args = [browser_exe]
                            if url:
                                args.append("--new-window")
                                args.append(url)
                            else:
                                args.append("--new-window")
                            
                            subprocess.Popen(args, cwd=cwd)
                            launched_count += 1
                            time.sleep(2.0) # Browsers need significant time to appear
                    elif cmdline:
                        subprocess.Popen(cmdline, cwd=cwd)
                        launched_count += 1
                        time.sleep(0.5) 
                except Exception as e:
                    print(f"Launch failed: {e}")

        if launched_count > 0:
            max_retries = 30
            target_count = initial_window_count + launched_count
            print(f"Waiting for windows... Initial: {initial_window_count}, Launched: {launched_count}, Target: {target_count}")
            
            for _ in range(max_retries):
                temp_windows = self.get_target_windows(detailed_scan=False)
                if len(temp_windows) >= target_count:
                    break
                time.sleep(0.5)

        # 2. Match & Restore
        current_windows = self.get_target_windows(detailed_scan=True)
        used_hwnds = set()
        
        matches = [] 

        # Match logic remains Top-to-Bottom to pairing correctness (prioritize main windows)
        for saved in saved_windows:
            saved_key = saved["title_pattern"]
            saved_url = saved.get("url")
            saved_folder = saved.get("folder_path")
            saved_cmdline = saved.get("cmdline")
            saved_exe = saved_cmdline[0] if saved_cmdline and len(saved_cmdline) > 0 else None
            
            best_match = None
            soft_match = None # Catch-all for same app but wrong/missing URL (e.g. Loading...)
            
            for current in current_windows:
                if current["hwnd"] in used_hwnds: continue
                
                # --- IDENTIFICATION CHECK ---
                # Default: Filter by Title Key
                is_candidate = (current["target_key"] == saved_key)
                
                # Better: Filter by Executable if available
                current_cmdline = current.get("cmdline")
                current_exe = current_cmdline[0] if current_cmdline and len(current_cmdline) > 0 else None
                
                if saved_exe and current_exe:
                    # If both have exe paths, use that for identification (ignore title changes)
                    if saved_exe.lower() == current_exe.lower():
                        is_candidate = True
                    else:
                        is_candidate = False
                
                if not is_candidate: continue
                
                # --- CONTENT MATCHING ---
                current_url = current.get("url")
                current_folder = current.get("folder_path")
                
                # Priority 1: Exact Folder Path
                if saved_folder and current_folder:
                    if saved_folder == current_folder:
                        best_match = current
                        break # Perfect match found
                        
                # Priority 2: URL Match (Partial)
                elif saved_url and current_url:
                    if saved_url in current_url or current_url in saved_url:
                        best_match = current
                        break # Good match found
                
                # Priority 3: No specific content needed (Calculator, Notepad...)
                elif not saved_url and not saved_folder:
                    if not best_match: best_match = current
                
                # Capture Soft Match (Same App, but URL mismatch/loading)
                # We only take this if we haven't found a better one yet
                if not soft_match: 
                    soft_match = current
            
            # If no perfect content match, fallback to soft match (e.g. Chrome "New Tab" vs "Google")
            if not best_match and soft_match:
                best_match = soft_match
            
            if best_match:
                used_hwnds.add(best_match["hwnd"])
                matches.append((saved, best_match))

        # 3. Apply Positions (Bottom-to-Top Stack Rebuild)
        # We iterate in REVERSE of saved order (Bottom first).
        # We force each window to HWND_TOP.
        # Bottom -> Top (Z=0)
        # Middle -> Top (Z=0, covers Bottom)
        # Top -> Top (Z=0, covers Middle)
        
        for saved, current in reversed(matches):
            hwnd = current["hwnd"]
            rect = saved["rect"]
            show_cmd = saved.get("show_cmd", win32con.SW_SHOWNORMAL)
            
            try:
                # Ensure visibility first
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                
                if show_cmd == win32con.SW_SHOWMINIMIZED or show_cmd == win32con.SW_SHOWMAXIMIZED:
                     flags = win32con.WPF_SETMINPOSITION if show_cmd == win32con.SW_SHOWMINIMIZED else 0
                     placement = (0, show_cmd, (-1, -1), (-1, -1), tuple(rect))
                     win32gui.SetWindowPlacement(hwnd, placement)
                     
                     # If maximized, we STILL want it on top of the previous one.
                     # SetWindowPos affects Z-order even for max windows.
                     win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                else:
                    x, y, r, b = rect
                    w = r - x
                    h = b - y
                    flags = win32con.SWP_SHOWWINDOW | win32con.SWP_FRAMECHANGED | win32con.SWP_NOOWNERZORDER
                    win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, w, h, flags)
                
                time.sleep(0.05) # Tiny delay to allow WM to process Z-order
            except Exception as e:
                print(f"Error moving window: {e}")

if __name__ == "__main__":
    import ctypes
    try:
        script_name = os.path.splitext(os.path.basename(__file__))[0]
        myappid = f'obat.{script_name}.v1' 
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
    app = WindowLayoutManager(root)
    root.mainloop()
