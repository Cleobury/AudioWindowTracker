# Audio Window Tracker

A Windows-native background service that tracks the position of audio-producing windows on your multi-monitor setup and automatically adjusts the left and right audio channels to create a 2D spatial / directional audio effect.

## Features
- **Horizontal Panning**: Automatically maps audio left/right balance based on the window's physical X coordinate.
- **System Tray Icon**: Runs silently in the background with a convenient taskbar icon.
- **Adjustable Intensity**: Choose between Low, Medium, or High panning strength.
- **Vertical Distance Attenuation**: Simulates Y-axis distance by subtly lowering overall volume.
- **Exponential Smoothing**: Prevents jarring audio jumps when dragging windows.
- **Minimized & Fullscreen Handling**: Automatically centers audio when windows are minimized or fullscreen.

## Installation & Setup

### For Most Users (Recommended)
1. Download the latest `AudioWindowTracker.exe`.
2. Move it to a folder of your choice (e.g., `C:\Apps\AudioTracker`).
3. Double-click to run. It will appear as an icon in your system tray (bottom right).

### Enabling Autostart
To have the app start automatically with Windows:
1. Open the `scripts` folder.
2. Right-click `enable_autostart.ps1` and select **Run with PowerShell**.

## How to Use
1. **Launch**: Open `AudioWindowTracker.exe`.
2. **Settings**: Right-click the tray icon to change the **Intensity**:
   - **Low**: Subtle effect (minimum 50% volume on opposite side).
   - **Medium**: Balanced (minimum 20% volume).
   - **High**: Maximum effect (minimum 0% volume).
3. **Exit**: Right-click the tray icon and select **Quit**.

## Note on Browser Audio
Modern browsers (Chrome, Edge, Brave, etc.) mix all tab audio into a single Windows Core Audio session. The tracker will target the first visible window it finds for that browser executable. To track multiple streams independently, use different browsers (e.g., Brave on the left, Edge on the right).

## Requirements
- Windows 10/11
- Active audio sessions (the app must be making sound to be panned)
