import uiautomation as auto
import win32gui

def get_browser_urls():
    print("Scanning for browsers...")
    
    def enum_handler(hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if "Google Chrome" in title or "Mozilla Firefox" in title:
                print(f"\nFound Target: {title}")
                
                try:
                    window = auto.ControlFromHandle(hwnd)
                    
                    # Chrome / Edge often use 'Address and search bar' or simple Edit controls
                    # Firefox uses 'Search with Google or enter address'
                    
                    # Search specifically for Edit controls
                    # We increase depth slightly and look for typical names
                    found_url = None
                    
                    # Strategy 1: Look for distinct names
                    params = [
                        {"ControlType": auto.ControlType.EditControl, "Name": "Address and search bar"}, # Chrome/Edge
                        {"ControlType": auto.ControlType.EditControl, "Name": "Search with Google or enter address"}, # Firefox
                        {"ControlType": auto.ControlType.EditControl, "Name": "Address"}, # Generic
                    ]
                    
                    for p in params:
                        control = window.Control(**p)
                        if control.Exists(maxSearchSeconds=0.5):
                             # Try getting value via ValuePattern
                            try:
                                url = control.GetValuePattern().Value
                                if url:
                                    print(f"  -> Found URL (Strategy 1 - {p['Name']}): {url}")
                                    found_url = url
                                    break
                            except:
                                pass
                                
                    if not found_url:
                        # Strategy 2: Just find ANY EditControl that looks like a URL
                        # This is riskier but covers variations
                        edits = window.GetChildren() # Get top level children first to avoid deep recursion
                        # Actually, let's just use GetFirstChildControl with condition
                        
                        edit = window.EditControl(searchDepth=5)
                        if edit.Exists(maxSearchSeconds=1):
                             try:
                                val = edit.GetValuePattern().Value
                                print(f"  -> Found URL (Strategy 2 - Generic Edit): {val}")
                             except:
                                print(f"  -> Edit found but no ValuePattern: {edit.Name}")
                                
                except Exception as e:
                    print(f"  -> Error accessing: {e}")

    win32gui.EnumWindows(enum_handler, None)

if __name__ == "__main__":
    get_browser_urls()
