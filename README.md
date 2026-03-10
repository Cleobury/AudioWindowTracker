# Audio Window Tracker

A Python script that tracks the position of audio-producing windows on your multi-monitor setup and automatically adjusts the left and right audio channels to create a 2D spatial / directional audio effect.

## Features
- **Horizontal Panning**: Automatically maps audio left/right balance based on the window's physical X coordinate across all monitors.
- **Vertical Distance Attenuation**: Simulates Y-axis distance by subtly lowering the overall volume (up to 20%) as the window moves further from the vertical center of the screens.
- **Exponential Smoothing**: Prevents jarring audio jumps when dragging windows quickly by smoothing the volume transitions.
- **Minimized & Fullscreen Handling**: Automatically snaps audio back to the center (100% L, 100% R) when a window is minimized or made fullscreen.

> **Note on Browser Audio**: Due to how modern browsers (Chrome, Edge, Brave, etc.) mix all tab audio into a single Windows Core Audio session, this script cannot split audio between multiple windows of the *same* browser simultaneously. It will target the first visible window it finds for that browser executable. To track multiple streams independently, use different browsers (e.g., Brave on the left, Edge on the right).

## Requirements
- Windows OS
- Python 3.8+

## Installation
1. Clone or download this repository.
2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage
Run the script from your terminal:
```bash
python main.py
```
Leave the terminal open while you play audio from an application and drag its window around your screens!
