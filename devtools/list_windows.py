import win32gui

def enum_windows_callback(hwnd, window_list):
    if win32gui.IsWindowVisible(hwnd):
        title = win32gui.GetWindowText(hwnd)
        if title:
            window_list.append((hwnd, title))

windows = []
win32gui.EnumWindows(enum_windows_callback, windows)

print("Visible Windows:")
for hwnd, title in windows:
    print(f"{hwnd}: {title}")
