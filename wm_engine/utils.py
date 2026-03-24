import win32api
from difflib import SequenceMatcher

def calculate_similarity(s1, s2):
    """ Returns a similarity score between 0.0 and 1.0 """
    return SequenceMatcher(None, s1, s2).ratio()

def clean_title(title):
    if not title: return ""
    t = title.lower()
    # Strip common browser suffixes
    suffixes = [
        " - google chrome", 
        " — mozilla firefox", 
        " - mozilla firefox", 
        " - microsoft edge", 
        " — navigation privée de mozilla firefox"
    ]
    for s in suffixes:
        if t.endswith(s):
            return title[:-len(s)].strip()
    return title

def normalize_url(url):
    if not url: return None
    if url.startswith("localhost") and not url.startswith("http"):
        return "http://" + url
    if not url.startswith("http") and not url.startswith("file") and "://" not in url:
        return "https://" + url
    return url

def capture_display_profile():
    try:
        monitors = []
        for index, monitor in enumerate(win32api.EnumDisplayMonitors()):
            info = win32api.GetMonitorInfo(monitor[0])
            rect = list(info["Monitor"])
            work_rect = list(info.get("Work", info["Monitor"]))
            monitors.append({
                "device": info.get("Device") or f"MONITOR_{index}",
                "rect": rect,
                "work_rect": work_rect,
                "is_primary": bool(info.get("Flags", 0) & 1),
                "profile_index": index,
            })

        monitors.sort(key=lambda m: (m["device"].lower(), m["rect"][0], m["rect"][1]))
        for index, monitor in enumerate(monitors):
            monitor["profile_index"] = index

        if monitors:
            left = min(m["rect"][0] for m in monitors)
            top = min(m["rect"][1] for m in monitors)
            right = max(m["rect"][2] for m in monitors)
            bottom = max(m["rect"][3] for m in monitors)
            primary_device = next((m["device"] for m in monitors if m["is_primary"]), monitors[0]["device"])
        else:
            left = top = right = bottom = 0
            primary_device = None

        return {
            "virtual_rect": [left, top, right, bottom],
            "primary_device": primary_device,
            "monitors": monitors,
        }
    except Exception:
        return {
            "virtual_rect": [0, 0, 0, 0],
            "primary_device": None,
            "monitors": [],
        }

def display_profiles_differ(saved_profile, current_profile):
    if not saved_profile or not current_profile:
        return False

    def canonical(profile):
        monitors = []
        for monitor in profile.get("monitors", []):
            monitors.append((
                str(monitor.get("device", "")).lower(),
                tuple(monitor.get("rect", [])),
                bool(monitor.get("is_primary", False)),
            ))
        return tuple(monitors)

    return canonical(saved_profile) != canonical(current_profile)

def get_rect_display_context(rect, display_profile=None):
    if not rect or len(rect) != 4:
        return None

    profile = display_profile or capture_display_profile()
    monitor = _find_monitor_for_rect(rect, profile.get("monitors", []))
    if not monitor:
        return None

    return {
        "device": monitor.get("device"),
        "monitor_rect": list(monitor.get("rect", [])),
        "work_rect": list(monitor.get("work_rect", [])),
        "is_primary": monitor.get("is_primary", False),
        "profile_index": monitor.get("profile_index"),
    }

def adapt_rect_to_display(rect, saved_display, current_profile):
    if not rect or len(rect) != 4 or not saved_display or not current_profile:
        return rect

    old_monitor_rect = saved_display.get("monitor_rect")
    if not old_monitor_rect or len(old_monitor_rect) != 4:
        return rect

    target_monitor = _find_target_monitor(saved_display, current_profile.get("monitors", []))
    if not target_monitor:
        return rect

    old_left, old_top, old_right, old_bottom = old_monitor_rect
    new_left, new_top, new_right, new_bottom = target_monitor["rect"]
    old_width = max(1, old_right - old_left)
    old_height = max(1, old_bottom - old_top)
    new_width = max(1, new_right - new_left)
    new_height = max(1, new_bottom - new_top)

    rel_left = (rect[0] - old_left) / old_width
    rel_top = (rect[1] - old_top) / old_height
    rel_right = (rect[2] - old_left) / old_width
    rel_bottom = (rect[3] - old_top) / old_height

    adapted = [
        round(new_left + rel_left * new_width),
        round(new_top + rel_top * new_height),
        round(new_left + rel_right * new_width),
        round(new_top + rel_bottom * new_height),
    ]
    return _clamp_rect_to_monitor(adapted, target_monitor["rect"])

def _find_monitor_for_rect(rect, monitors):
    if not monitors:
        return None

    best_monitor = None
    best_area = -1
    for monitor in monitors:
        area = _intersection_area(rect, monitor["rect"])
        if area > best_area:
            best_area = area
            best_monitor = monitor

    if best_monitor and best_area > 0:
        return best_monitor

    cx = (rect[0] + rect[2]) / 2
    cy = (rect[1] + rect[3]) / 2
    return min(
        monitors,
        key=lambda monitor: _distance_sq_to_rect_center(cx, cy, monitor["rect"]),
    )

def _find_target_monitor(saved_display, monitors):
    if not monitors:
        return None

    saved_device = str(saved_display.get("device", "")).lower()
    if saved_device:
        for monitor in monitors:
            if str(monitor.get("device", "")).lower() == saved_device:
                return monitor

    saved_index = saved_display.get("profile_index")
    if isinstance(saved_index, int) and 0 <= saved_index < len(monitors):
        return monitors[saved_index]

    saved_rect = saved_display.get("monitor_rect") or [0, 0, 1, 1]
    saved_width = max(1, saved_rect[2] - saved_rect[0])
    saved_height = max(1, saved_rect[3] - saved_rect[1])
    return min(
        monitors,
        key=lambda monitor: (
            abs((monitor["rect"][2] - monitor["rect"][0]) - saved_width)
            + abs((monitor["rect"][3] - monitor["rect"][1]) - saved_height),
            0 if monitor.get("is_primary") == saved_display.get("is_primary", False) else 1,
            monitor.get("profile_index", 0),
        ),
    )

def _intersection_area(rect_a, rect_b):
    left = max(rect_a[0], rect_b[0])
    top = max(rect_a[1], rect_b[1])
    right = min(rect_a[2], rect_b[2])
    bottom = min(rect_a[3], rect_b[3])
    if right <= left or bottom <= top:
        return 0
    return (right - left) * (bottom - top)

def _distance_sq_to_rect_center(x, y, rect):
    cx = (rect[0] + rect[2]) / 2
    cy = (rect[1] + rect[3]) / 2
    dx = cx - x
    dy = cy - y
    return dx * dx + dy * dy

def _clamp_rect_to_monitor(rect, monitor_rect):
    left, top, right, bottom = rect
    width = max(50, right - left)
    height = max(50, bottom - top)
    mon_left, mon_top, mon_right, mon_bottom = monitor_rect
    mon_width = max(1, mon_right - mon_left)
    mon_height = max(1, mon_bottom - mon_top)

    width = min(width, mon_width)
    height = min(height, mon_height)
    left = min(max(left, mon_left), mon_right - width)
    top = min(max(top, mon_top), mon_bottom - height)
    return [int(left), int(top), int(left + width), int(top + height)]

def ensure_rect_on_screen(rect):
    """
    Ensures the given window rect is visible on at least one monitor.
    If not, moves it to the primary monitor.
    rect = [x, y, x2, y2]
    """
    try:
        x, y, r, b = rect
        w = r - x
        h = b - y
        
        # Center point of the window
        cx = x + (w // 2)
        cy = y + (h // 2)
        
        monitors = win32api.EnumDisplayMonitors()
        for monitor in monitors:
            # monitor[2] is the rect (left, top, right, bottom)
            mx, my, mr, mb = monitor[2]
            if mx <= cx <= mr and my <= cy <= mb:
                return rect # Center is inside a monitor, executed as is.

        # If we are here, the window is off-screen.
        # Reset to primary monitor (0,0) with some padding
        print(f"Window {rect} is off-screen. Resetting to primary monitor.")
        return [50, 50, 50 + w, 50 + h]
        
    except Exception as e:
        print(f"Error checking screen bounds: {e}")
        return rect
