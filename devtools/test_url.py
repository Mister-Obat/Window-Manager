import uiautomation as auto
import time

def get_browser_url():
    print("Switch to a browser window in 3 seconds...")
    time.sleep(3)
    
    hwnd = auto.GetForegroundWindow()
    window = auto.ControlFromHandle(hwnd)
    print(f"Window: {window.Name}")
    
    # Try finding the address bar
    # Chrome/Edge: Edit control with specific names or access keys
    # Firefox: Edit control "Search with Google or enter address"
    
    # Generic approach: Find the first Edit control in the upper part of the window (Toolbar)
    
    # Limit search depth for speed
    edit = window.EditControl(searchDepth=7)
    
    if edit.Exists(maxSearchSeconds=1):
        print(f"Found Edit Control: {edit.Name}")
        try:
            # ValuePattern is often used for URL bars
            url = edit.GetValuePattern().Value
            print(f"URL: {url}")
        except:
            print("No ValuePattern, trying Name...")
            print(f"Name as Value: {edit.Name}")
    else:
        print("Address bar not found via generic EditControl search.")
        
        # Fallback: traverse tree to find something creating "Address"
        # This is slower/harder.

if __name__ == "__main__":
    get_browser_url()
