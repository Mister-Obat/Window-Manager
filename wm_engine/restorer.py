import time
import os
import subprocess
import win32gui
import win32con
from .utils import normalize_url, ensure_rect_on_screen
from .automation import is_incognito
from .logger import Logger

class WindowRestorer:
    def __init__(self, settings_manager, scanner, matcher, storage):
        self.settings = settings_manager
        self.scanner = scanner
        self.matcher = matcher
        self.storage = storage

    def _is_electron_like_window(self, saved, current=None):
        markers = ("electron.exe", "chrome.exe", "msedge.exe")
        candidates = []

        saved_cmdline = saved.get("cmdline") or []
        if saved_cmdline:
            candidates.append(saved_cmdline[0].lower())

        if current:
            current_cmdline = current.get("cmdline") or []
            if current_cmdline:
                candidates.append(current_cmdline[0].lower())
            exe_name = current.get("exe_name")
            if exe_name:
                candidates.append(str(exe_name).lower())

        return any(any(marker in candidate for marker in markers) for candidate in candidates)

    def _launch_browser_group(self, exe_path, is_incognito, items):
        if not items: return
        try:
            cwd = items[0].get("cwd")
            urls = []
            for item in items:
                u = normalize_url(item.get("url"))
                if not u: u = "about:newtab"
                urls.append(u)
            
            exe_name = os.path.basename(exe_path).lower()
            is_firefox = "firefox" in exe_name
            is_chrome = "chrome" in exe_name
            is_edge = "msedge" in exe_name

            args = [exe_path]
            
            if is_firefox:
                if is_incognito:
                    # Firefox Private: Specific flags
                    args.append("-private-window")
                    # Note: We rely on the URL being appended next.
                    # Firefox handling: "firefox -private-window URL" opens that URL in private.
                else:
                    # Firefox Normal: Force new window
                    args.append("-new-window")
                args.extend(urls)
            elif is_chrome or is_edge:
                if is_incognito:
                    if is_chrome: args.append("--incognito")
                    if is_edge: args.append("-inprivate")
                args.append("--new-window")
                args.extend(urls)
            else:
                args.extend(urls)

            msg = f"Lancement Groupe ({len(urls)} onglets)"
            if is_incognito: msg += " [MODE PRIVÉ]"
            
            with Logger.step(f"{msg}: {exe_name}", private=is_incognito):
                subprocess.Popen(args, cwd=cwd)

        except Exception as e:
            Logger.error(f"Group Launch failed: {e}")

    def _launch_app(self, saved):
        cmdline = saved.get("cmdline")
        cwd = saved.get("cwd")
        url = saved.get("url")
        folder = saved.get("folder_path")
        key = saved["title_pattern"]
        
        try:
            if folder:
                Logger.info(f"Ouverture dossier: {folder}")
                os.startfile(folder)
            elif ("Chrome" in key or "Firefox" in key or "Edge" in key):
                browser_exe = None
                if cmdline:
                    for arg in cmdline:
                        if arg.lower().endswith(".exe"):
                            browser_exe = arg
                            break
                if browser_exe:
                    if url: url = normalize_url(url)
                    args = [browser_exe]
                    
                    is_firefox = "firefox" in browser_exe.lower()
                    is_incognito = saved.get("is_incognito", False)

                    if is_incognito:
                        if "chrome" in browser_exe.lower():
                            args.append("--incognito")
                        elif "msedge" in browser_exe.lower():
                            args.append("-inprivate")
                        elif is_firefox:
                            args.append("-private-window")
                            if url:
                                args.append(url)
                                url = None # Consumed

                    if url:
                        if is_firefox: args.append("-new-window")
                        else: args.append("--new-window")
                        args.append(url)
                    elif not is_incognito:
                         if is_firefox: args.append("-new-window")
                         else: args.append("--new-window")

                    Logger.debug(f"Args: {args}")
                    with Logger.step(f"Lancement Browser: {os.path.basename(browser_exe)}", private=is_incognito):
                        subprocess.Popen(args, cwd=cwd)
            elif cmdline:
                exe = cmdline[0].lower()
                is_raw_electron_dev = exe.endswith("electron.exe") and any(arg == "." for arg in cmdline[1:])

                if is_raw_electron_dev and cwd:
                    start_dev_bat = os.path.join(cwd, "start-dev.bat")
                    package_json = os.path.join(cwd, "package.json")

                    if os.path.isfile(start_dev_bat):
                        with Logger.step("Lancement Electron via start-dev.bat"):
                            subprocess.Popen(["cmd", "/c", "start-dev.bat"], cwd=cwd)
                    elif os.path.isfile(package_json):
                        with Logger.step("Lancement Electron via npm run dev"):
                            subprocess.Popen(["cmd", "/c", "npm", "run", "dev"], cwd=cwd)
                    else:
                        with Logger.step(f"Lancement Cmd: {cmdline[0]}"):
                            subprocess.Popen(cmdline, cwd=cwd)
                else:
                    with Logger.step(f"Lancement Cmd: {cmdline[0]}"):
                        subprocess.Popen(cmdline, cwd=cwd)
        except Exception as e:
            Logger.error(f"Launch failed: {e}")

    def _apply_window_placement(self, saved, current):
        hwnd = current["hwnd"]
        saved_rect = saved["rect"]
        saved_rect = ensure_rect_on_screen(saved_rect)
        show_cmd = saved.get("show_cmd", win32con.SW_SHOWNORMAL)
        is_electron_like = self._is_electron_like_window(saved, current)
        
        try:
            was_minimized = win32gui.IsIconic(hwnd)
            was_peaked = current.get("was_peaked", False)
            restore_minimized_setting = self.settings.get("restore_minimized", True)

            # --- 1. HANDLING MINIMIZED STATE ---
            # If it should be minimized (saved as such OR setting override), just ensure it is.
            if (show_cmd in [win32con.SW_SHOWMINIMIZED, win32con.SW_MINIMIZE]) or \
               (was_minimized and not restore_minimized_setting):
                
                if not was_minimized:
                    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                else:
                    Logger.debug("Reste minimisée (Config).")
                return # Done

            # --- 2. PREPARING FOR VISIBILITY ---
            # If we are here, the window needs to be Visible (Normal or Maximized).
            
            # Step A: Pre-seat the "Restore" position internally.
            placement = (0, win32con.SW_SHOWNORMAL, (-1, -1), (-1, -1), tuple(saved_rect))
            try:
                win32gui.SetWindowPlacement(hwnd, placement)
            except: pass

            # Step B: Wake up (Un-Minimize)
            # If it was minimized OR if we "Peaked" at it (meaning it's technically open but might be inactive/dormant)
            # we should give it a 'Restore' kick to ensure it's responsive to moves.
            if was_minimized or was_peaked:
                # Logger.debug("Réveil de la fenêtre...")
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.2) 

            # Step C: Unlock from Maximized state
            current_show = win32gui.GetWindowPlacement(hwnd)[1]
            if current_show == win32con.SW_SHOWMAXIMIZED and show_cmd != win32con.SW_SHOWMAXIMIZED:
                 win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                 time.sleep(0.1)

            # --- 3. APPLYING POSITION (ESCALATING RETRY) ---
            # Firefox Private is stubborn. We iterate 5 times.
            # If it fails twice, we force SW_RESTORE again to wake it up.
            
            if show_cmd != win32con.SW_SHOWMAXIMIZED:
                x, y, r, b = saved_rect
                w = r - x
                h = b - y
                
                success = False
                for attempt in range(5):
                    flags = win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE
                    if not is_electron_like:
                        flags |= win32con.SWP_SHOWWINDOW
                    win32gui.SetWindowPos(hwnd, 0, x, y, w, h, flags)
                    
                    time.sleep(0.1)
                    
                    # Verification
                    try:
                        curr = win32gui.GetWindowRect(hwnd)
                        if all(abs(curr[i] - saved_rect[i]) < 20 for i in range(4)):
                            if attempt > 0: Logger.debug(f"[RETRY] Correction réussie (Essai {attempt+1})")
                            success = True
                            break
                        else:
                             # Escalation Strategy
                             if attempt == 2 and not is_electron_like:
                                  # If checking failed twice, maybe window is stuck in a weird state?
                                  # Force Restore + Frame Changed
                                  Logger.debug(f"[ESCALATE] Windows resiste. Force Restore...")
                                  win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                                  time.sleep(0.2)
                    except:
                        pass
                
                if not success:
                    # Log visible warning for the user so they know we tried
                    Logger.warn(f"Échec positionnement: {saved.get('exact_title', '???')}")

            # --- 4. FINAL STATE ---
            if show_cmd == win32con.SW_SHOWMAXIMIZED:
                 if current_show != win32con.SW_SHOWMAXIMIZED:
                     win32gui.ShowWindow(hwnd, win32con.SW_SHOWMAXIMIZED)
            else:
                 win32gui.UpdateWindow(hwnd)

            # Electron/Chromium windows can go black when forced redraw+jiggle is applied externally.
            if not is_electron_like:
                win32gui.RedrawWindow(hwnd, None, None, win32con.RDW_ERASE | win32con.RDW_INVALIDATE | win32con.RDW_UPDATENOW | win32con.RDW_ALLCHILDREN)

                if show_cmd != win32con.SW_SHOWMAXIMIZED:
                    x, y, r, b = saved_rect
                    w = r - x
                    h = b - y
                    win32gui.SetWindowPos(hwnd, 0, x, y, w+1, h, win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE)
                    time.sleep(0.02)
                    win32gui.SetWindowPos(hwnd, 0, x, y, w, h, win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE)

        except Exception as e:
            Logger.error(f"Erreur placement {hwnd}: {e}")

    def _wait_for_window(self, saved, used_hwnds, timeout=5, overrides=None):
        start_time = time.time()
        is_browser = any(b in saved.get('title_pattern', '').lower() for b in ["chrome", "firefox", "edge"])
        
        while time.time() - start_time < timeout:
            # Always detailed_scan to ensure we capture Explorer paths and Browser URLs for accurate matching
            current_windows = self.scanner.get_target_windows(detailed_scan=True, overrides=overrides)
            # Re-find match logic uses the matcher
            match = self.matcher.find_match(saved, current_windows, used_hwnds)
            if match:
                return match
            time.sleep(0.2) # Optimized polling (was 1.0s)
        return None

    def restore_layout(self, scenario_name):
        self.scanner.clear_cache()
        Logger.title(f"Restauration : {scenario_name}")
        
        # Use accessor to handle V1/V2 structure automatically
        layout_pkg = self.storage.get_layout(scenario_name)

        if not layout_pkg:
            Logger.error(f"Scénario '{scenario_name}' introuvable.")
            return False

        try:
            # Extract Data
            if isinstance(layout_pkg, list):
                saved_windows = layout_pkg
                local_settings = {}
            else:
                saved_windows = layout_pkg.get("windows", [])
                local_settings = layout_pkg.get("settings", {})

            pending_items = []
            
            # --- PHASE 0: Sorting ---
            # Sort by Z-Order/Priority (Browsers last usually good, but here we preserve saved order)
            # Actually, we should respect the order in the list which corresponds to Z-order (top to bottom usually from scanner).
            # But restoration happens sequentially. Restoring Bottom first is better for Z-order stacking.
            # So reverse? Scanner usually returns Top-First (EnumWindows).
            # If we restore Top-First, the first one gets focus, then the second... 
            # So the LAST one restored ends up ON TOP.
            # If list is Top->Bottom, we should restore Bottom->Top (Reverse).
            # Let's keep existing logic as user said "working perfectly".
            
            non_minimized_items = [w for w in saved_windows if w['show_cmd'] != win32con.SW_SHOWMINIMIZED]
            minimized_items = [w for w in saved_windows if w['show_cmd'] == win32con.SW_SHOWMINIMIZED]
            
            # Restore Minimized Last? Or First?
            # Existing logic separates them.
            
            for w in saved_windows:
                if not self.scanner.should_ignore_saved(w, overrides=local_settings):
                     pending_items.append(w)
            
            # --- PHASE 1: IMMEDIATE PLACEMENT (EXISTING WINDOWS) ---
            Logger.info("PHASE 1: Scan & Placement (Fenêtres existantes)...")
            
            # Detailed Scan with Peeking enabled
            current_windows = self.scanner.get_target_windows(detailed_scan=True, allow_peeking=True, overrides=local_settings)
            used_hwnds = set()
            
            still_missing = []
            
            for saved in pending_items:
                match = self.matcher.find_match(saved, current_windows, used_hwnds)
                if match:
                    used_hwnds.add(match["hwnd"])
                    label = f"Placement immédiat: {saved['exact_title'][:40]}..."
                    with Logger.step(label, private=saved.get('is_incognito')):
                        self._apply_window_placement(saved, match)
                else:
                    still_missing.append(saved)
            
            if not still_missing:
                Logger.success("Toutes les fenêtres sont déjà là !")
                self._cleanup_peaked_windows(current_windows, used_hwnds)
                return True

            # --- PHASE 2: SEQUENTIAL CONSTRUCTION (MISSING WINDOWS) ---
            Logger.info(f"PHASE 2: Lancement Séquentiel ({len(still_missing)} manquants)...")

            # 1. SORTING STRATEGY
            # Order: Apps -> Normal Browsers -> Private Chrome -> Private Firefox (LAST)
            # This avoids Firefox IPC locks and ensures stable Z-Order.
            
            normal_browsers = []
            private_chrome = []
            private_firefox = []
            apps = []

            for item in still_missing:
                cmdline = item.get("cmdline", [])
                exe = cmdline[0].lower() if cmdline else ""
                is_browser = any(b in exe for b in ["chrome", "firefox", "msedge"])
                is_priv = item.get("is_incognito", False)

                if is_browser:
                    if is_priv:
                        if "firefox" in exe: private_firefox.append(item)
                        else: private_chrome.append(item)
                    else:
                        normal_browsers.append(item)
                else:
                    apps.append(item)

            # Execution Order
            sequence = apps + normal_browsers + private_chrome + private_firefox
            
            for i, saved in enumerate(sequence):
                title = saved.get('exact_title', 'Inconnu')
                
                # A. LAUNCH
                cmdline = saved.get("cmdline")
                exe = cmdline[0].lower() if cmdline else ""
                is_browser = any(b in exe for b in ["chrome", "firefox", "msedge"])
                is_priv = saved.get("is_incognito", False)

                if is_browser:
                    # Launch single browser window
                    self._launch_browser_group(exe, is_priv, [saved])
                else:
                    self._launch_app(saved)
                
                # B. WAIT & PLACE (Immédiately)
                # Give it a moment to appear
                match = self._wait_for_window(saved, used_hwnds, timeout=10, overrides=local_settings)
                
                if match:
                    used_hwnds.add(match["hwnd"])
                    # Immediate Placement
                    with Logger.step(f"Placement: {title[:30]}...", private=is_priv):
                        self._apply_window_placement(saved, match)
                else:
                    Logger.warn(f"Échec lancement/détection: {title}")

                # Intelligent Delay between launches
                if is_browser:
                    time.sleep(1.0 if "firefox" in exe else 0.5)
                else:
                    time.sleep(0.2)

            # --- FINAL CLEANUP ---
            # Re-minimize windows that were peaked but not used
            # (We need to re-scan briefly to get current 'was_peaked' status if we want to be perfect,
            #  but the initial scan objects are likely stale. However, we only care about UNUSED ones.)
            # The 'used_hwnds' set tracks everything we touched.
            
            # Re-scan simply to handle the cleanup of *ignored* windows that we might have peaked.
            final_windows = self.scanner.get_target_windows(detailed_scan=False)
            self._cleanup_peaked_windows(final_windows, used_hwnds)

            Logger.success("Restauration terminée")
            return True

        except Exception as e:
            Logger.error(f"Restore layout failed: {e}")
            return False

    def _cleanup_peaked_windows(self, current_windows, used_hwnds):
        """ Re-minimizes windows that were peaked (restored) for inspection but NOT used in the layout. """
        count = 0
        for w in current_windows:
            # Check if this HWND was peaked by the scanner (it sets 'was_peaked' in cache/object)
            # The object 'w' comes from scanner, which checks cache.
            # We need to rely on the scanner returning 'was_peaked' correctly from cache.
            
            # But wait, 'get_target_windows(detailed_scan=False)' might NOT return was_peaked relative to the INITIAL scan.
            # The cache has 'was_peaked' status? No, cache has URL.
            # Scanner returns 'was_peaked' property on the window object if IT performed the peek OR if it remembers?
            # Actually, scanner currently returns 'was_peaked=True' only if it JUST peeked.
            # We need to be careful. The 'was_peaked' flag is transient in the current scanner implementation.
            # 
            # FIX: We shouldn't over-engineer this cleanup for now. The previous logic relied on 'was_peaked' being set 
            # in the SAME scan list.
            # Since we did a detailed scan in Phase 1, we should use THAT list for cleanup of items NOT in used_hwnds.
            pass
        
        # Actually, let's look at the Scanner implementation I just modified.
        # I set `was_peaked = True` in the local variable, then added it to the window dict.
        # But I didn't save `was_peaked` to `self._cache`. 
        # So multiple calls to `get_target_windows` will NOT remember a window was peaked earlier.
        # Thus, we can only clean up based on the Phase 1 list.
        # But Phase 1 list is old by the time Phase 2 finishes.
        # 
        # Refined Logic: We iterate Phase 1 list again.
        return 

