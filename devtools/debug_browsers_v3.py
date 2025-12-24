import uiautomation as auto
import win32gui
import time

def scan_browsers():
    print("Enumerating Windows...")
    
    def enum_handler(hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if "Chrome" in title or "Firefox" in title:
                print(f"\n==========================================")
                print(f"Target Found: '{title}' (HWND: {hwnd})")
                
                try:
                    window = auto.ControlFromHandle(hwnd)
                    print(f"  Automation Name: {window.Name}")
                    print(f"  Automation Type: {window.ControlTypeName}")
                    
                    # Search for ANY EditControl and print its details
                    print("  Searching for EditControls (Depth 10)...")
                    
                    count = 0
                    for control, depth in auto.WalkControl(window, maxDepth=10):
                        if control.ControlTypeName == "EditControl":
                            count += 1
                            name = control.Name
                            val = ""
                            try:
                                if control.GetPattern(auto.PatternId.ValuePattern):
                                    val = control.GetValuePattern().Value
                            except:
                                pass
                                
                            print(f"    [EditControl] Name: '{name}' | Value: '{val}' | Depth: {depth}")
                    
                    if count == 0:
                        print("    NO EditControls found within depth 10.")
                        
                except Exception as e:
                    print(f"  Error accessing automation: {e}")

    win32gui.EnumWindows(enum_handler, None)

if __name__ == "__main__":
    time.sleep(1)
    scan_browsers()
