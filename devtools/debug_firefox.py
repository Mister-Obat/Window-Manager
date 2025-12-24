import uiautomation as auto
import win32gui
import time

def inspect_firefox():
    print("Switch to Firefox in 3 seconds...")
    time.sleep(3)
    
    hwnd = auto.GetForegroundWindow()
    window = auto.ControlFromHandle(hwnd)
    print(f"Inspecting Window: {window.Name} ({window.ControlTypeName})")

    print("\n--- Searching for Edit Controls (Depth 15) ---")
    
    count = 0
    # Walk tree
    for control, depth in auto.WalkControl(window, maxDepth=15):
        if control.ControlTypeName == "EditControl":
            print(f"[{depth}] Name: '{control.Name}'")
            try:
                if control.GetPattern(auto.PatternId.ValuePattern):
                    val = control.GetValuePattern().Value
                    print(f"    -> Value: '{val}'")
            except:
                pass
            count += 1
            
    print(f"\nTotal EditControls found: {count}")

if __name__ == "__main__":
    inspect_firefox()
