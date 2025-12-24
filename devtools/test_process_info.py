import win32gui
import win32process
import psutil

def get_process_info():
    def enum_handler(hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title:
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    proc = psutil.Process(pid)
                    cmdline = proc.cmdline()
                    print(f"Title: {title} | PID: {pid} | Cmd: {cmdline}")
                except Exception as e:
                    print(f"Title: {title} | Error: {e}")

    win32gui.EnumWindows(enum_handler, None)

if __name__ == "__main__":
    get_process_info()
