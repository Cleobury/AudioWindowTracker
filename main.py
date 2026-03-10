import pygetwindow as gw
from screeninfo import get_monitors
from pycaw.pycaw import AudioUtilities
import win32process
import psutil
import time
import threading
import pystray
from PIL import Image
import os
import sys
import logging
import pythoncom
import json

def get_total_width():
    monitors = get_monitors()
    total_width = sum(m.width for m in monitors)
    return total_width

def get_max_height():
    monitors = get_monitors()
    max_height = max(m.height for m in monitors)
    return max_height

def apply_directional_audio():
    total_width = get_total_width()
    max_height = get_max_height()
    
    # 1. Get executable names of processes playing audio and their corresponding sessions
    sessions = AudioUtilities.GetAllSessions()
    active_audio_sessions_by_exe = {}
    
    for session in sessions:
        if session.Process and session.State == 1: # Actively playing
            try:
                exe_name = session.Process.name().lower()
                if exe_name not in active_audio_sessions_by_exe:
                    active_audio_sessions_by_exe[exe_name] = []
                active_audio_sessions_by_exe[exe_name].append(session)
            except psutil.NoSuchProcess:
                pass
                
    if not active_audio_sessions_by_exe:
        return []

    results = []
    
    # 2. Get all visible windows
    windows = gw.getAllWindows()
    seen_exes = set()
    
    for win in windows:
        if not win.visible or not win.title:
            continue
            
        # Get PID for this window
        _, pid = win32process.GetWindowThreadProcessId(win._hWnd)
        
        try:
            win_exe = psutil.Process(pid).name().lower()
        except psutil.NoSuchProcess:
            win_exe = ""
            
        if win_exe in active_audio_sessions_by_exe and win_exe not in seen_exes:
            # We use the first visible window we find for a given executable to calculate panning
            seen_exes.add(win_exe)
            
            # Check if minimized. Windows moves minimized windows to -32000
            # Check if fullscreen/borderless. If the window is >= the total width of a single monitor, or covers the primary screen
            is_minimized = win.left <= -30000
            
            # We can check if it's fullscreen by seeing if its width and height match any monitor
            # or if it's exceptionally large
            is_fullscreen = False
            for m in get_monitors():
                if win.width >= m.width and win.height >= m.height:
                    is_fullscreen = True
                    break

            if is_minimized or is_fullscreen:
                panning = 0.0 # Center
                vertical_panning = 0.0 # Center
            else:
                window_centre_x = win.left + (win.width / 2)
                panning = (2 * window_centre_x - total_width) / total_width
                panning = max(-1.0, min(1.0, panning))
                
                window_centre_y = win.top + (win.height / 2)
                vertical_panning = (2 * window_centre_y - max_height) / max_height
                vertical_panning = max(-1.0, min(1.0, vertical_panning))

            # Calculate target L / R volumes based on X-axis panning
            # panning: -1.0 (Full Left) to 1.0 (Full Right)
            
            # Use dynamic min_vol from settings
            intensity = current_settings.get("intensity", "Medium")
            min_vol = 0.5 if intensity == "Low" else (0.2 if intensity == "Medium" else 0.0)
            range_vol = 1.0 - min_vol
            
            if panning > 0:
                target_left = 1.0 - (panning * range_vol)
                target_right = 1.0
            else:
                target_left = 1.0
                target_right = 1.0 + (panning * range_vol)
                
            # Apply Vertical Volume Attenuation
            # vertical_panning: 0.0 is center, -1.0 is top, 1.0 is bottom
            # Max drop is 20% (multiplier goes from 1.0 down to 0.8)
            y_distance = abs(vertical_panning)
            vertical_multiplier = 1.0 - (y_distance * 0.2)
            
            target_left *= vertical_multiplier
            target_right *= vertical_multiplier
                
            # Apply Smoothing (EMA)
            smoothing_factor = 0.3 # 0.0 to 1.0. Lower means slower/smoother transitions

            # Retrieve existing volumes instead of jumping straight to the target
            # If we don't know the current state, we assume the target to start.
            try:
                # Get the first session to check current volume
                first_session = active_audio_sessions_by_exe[win_exe][0]
                vol_ctrl = first_session.channelAudioVolume()
                if vol_ctrl.GetChannelCount() >= 2:
                    current_left = vol_ctrl.GetChannelVolume(0)
                    current_right = vol_ctrl.GetChannelVolume(1)
                else:
                    current_left, current_right = target_left, target_right
            except Exception:
                current_left, current_right = target_left, target_right

            # Calculate the new smoothed volume
            left_vol = current_left + smoothing_factor * (target_left - current_left)
            right_vol = current_right + smoothing_factor * (target_right - current_right)
            
            # Apply to all audio sessions assigned to this executable
            for session in active_audio_sessions_by_exe[win_exe]:
                try:
                    vol_ctrl = session.channelAudioVolume()
                    channels = vol_ctrl.GetChannelCount()
                    if channels >= 2:
                        vol_ctrl.SetChannelVolume(0, left_vol, None)
                        vol_ctrl.SetChannelVolume(1, right_vol, None)
                except Exception as e:
                    pass

            results.append({
                "title": win.title,
                "exe": win_exe,
                "panning": panning,
                "left_vol": left_vol,
                "right_vol": right_vol
            })
            
    return results


# Setup logging to a file in the app directory
log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "debug.log")
logging.basicConfig(
    filename=log_path,
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Settings Management
settings_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "settings.json")
current_settings = {"intensity": "Medium"}

def load_settings():
    global current_settings
    if os.path.exists(settings_path):
        try:
            with open(settings_path, 'r') as f:
                current_settings = json.load(f)
        except Exception as e:
            logging.error(f"Failed to load settings: {e}")

def save_settings():
    try:
        with open(settings_path, 'w') as f:
            json.dump(current_settings, f)
    except Exception as e:
        logging.error(f"Failed to save settings: {e}")

# Global control for the background thread
running = True
tracker_thread = None

def run_tracker():
    """Background loop for audio tracking."""
    global running
    load_settings()
    logging.info("Tracker thread started.")
    
    # Initialize COM for this thread
    try:
        pythoncom.CoInitialize()
        logging.info("COM initialized in tracker thread.")
    except Exception as e:
        logging.error(f"Failed to initialize COM: {e}")
        return

    while running:
        try:
            results = apply_directional_audio()
            if results:
                logging.debug(f"Panned {len(results)} windows.")
        except Exception as e:
            logging.error(f"Error in tracking loop: {e}", exc_info=True)
        time.sleep(0.1)
    
    pythoncom.CoUninitialize()
    logging.info("Tracker thread stopping.")

def on_quit(icon, item):
    """Callback to shut down the application."""
    global running
    running = False
    icon.stop()
    sys.exit(0)

def set_intensity(icon, item):
    global current_settings
    current_settings["intensity"] = item.text
    save_settings()
    logging.info(f"Intensity set to {item.text}")

def setup_tray():
    """Initializes and runs the system tray icon."""
    load_settings() # Load before building menu

    # Build icon
    icon_path = os.path.join(os.path.dirname(__file__), "app_icon.png")
    if not os.path.exists(icon_path):
        image = Image.new('RGB', (64, 64), (0, 120, 215))
    else:
        image = Image.open(icon_path)

    intensity_menu = pystray.Menu(
        pystray.MenuItem("Low", set_intensity, checked=lambda item: current_settings["intensity"] == "Low", radio=True),
        pystray.MenuItem("Medium", set_intensity, checked=lambda item: current_settings["intensity"] == "Medium", radio=True),
        pystray.MenuItem("High", set_intensity, checked=lambda item: current_settings["intensity"] == "High", radio=True)
    )

    menu = pystray.Menu(
        pystray.MenuItem("Audio Window Tracker", lambda: None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Intensity", intensity_menu),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Quit", on_quit)
    )

    icon = pystray.Icon("AudioTracker", image, "Audio Window Tracker", menu)
    
    # Start tracker in background thread
    global tracker_thread
    tracker_thread = threading.Thread(target=run_tracker, daemon=True)
    tracker_thread.start()

    icon.run()

if __name__ == "__main__":
    setup_tray()
