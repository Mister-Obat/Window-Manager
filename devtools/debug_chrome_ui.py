import uiautomation as auto
import win32gui
import time

def inspect_chrome():
    print("Switch to Chrome in 3 seconds...")
    time.sleep(3)
    
    hwnd = auto.GetForegroundWindow()
    window = auto.ControlFromHandle(hwnd)
    print(f"Inspecting Window: {window.Name}")

    # Walk the tree for the first few levels to find the address bar
    print("Walking UI Tree (Depth 10)...")
    
    def walk(control, depth=0):
        if depth > 10:
            return
        
        indent = "  " * depth
        try:
            name = control.Name
            ctrl_type = control.ControlTypeName
            val = ""
            
            try:
                # Try to get value if supported
                if control.GetPattern(auto.PatternId.ValuePattern):
                    val = f" [Value: {control.GetValuePattern().Value}]"
            except:
                pass
            
            print(f"{indent}{ctrl_type}: '{name}'{val}")
            
            # Simple heuristic to stop if we find something URL-like
            if "Address" in name or "Search" in name or "https://" in str(val) or "www." in str(val):
                print(f"{indent}*** POTENTIAL MATCH ***")

            for child in control.GetChildren():
                walk(child, depth + 1)
        except Exception as e:
            print(f"{indent}Error: {e}")

    walk(window, 0)

if __name__ == "__main__":
    inspect_chrome()
