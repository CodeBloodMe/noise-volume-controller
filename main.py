import sounddevice as sd
import numpy as np
import time
import threading
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import tkinter as tk
from tkinter import ttk

# --- Constants ---
SOUND_DEVICE_DURATION = 0.2
SAMPLE_RATE = 44100
MIN_VOLUME = 0.0
MAX_VOLUME = 1.0
UPDATE_INTERVAL = 0.15  # Update more frequently for smoother response
RESTING_NOISE_LEVEL = 0.005 # Assumed ambient noise level in a quiet room

# --- Global State ---
# This variable will track the last volume level set by this application.
# It's crucial for detecting when the user makes a change externally (e.g., with keyboard keys).
last_volume_set_by_app = -1.0
# A lock to prevent race conditions when accessing the above variable from multiple threads.
volume_lock = threading.Lock()

# --- Get mic noise ---
def get_noise_level(duration=SOUND_DEVICE_DURATION, sample_rate=SAMPLE_RATE):
    """Records audio and returns its Root Mean Square (RMS) to represent noise level."""
    try:
        audio = sd.rec(int(duration * sample_rate), samplerate=sample_rate, channels=1, dtype='float64')
        sd.wait()
        rms = np.sqrt(np.mean(audio**2))
        return rms
    except Exception as e:
        print(f"Error getting noise level: {e}")
        return 0.0

# --- Get Windows system volume interface ---
def get_volume_control():
    """Returns the master audio volume control interface for the default speakers."""
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    return cast(interface, POINTER(IAudioEndpointVolume))

# --- Auto volume loop ---
def monitor_noise(volume_control, get_sensitivity, get_toggle_state, get_listen_state, get_base_volume, update_base_volume_in_gui):
    """
    Monitors microphone noise and system volume to make intelligent adjustments.
    This function runs in a background thread.
    """
    global last_volume_set_by_app
    print("🎧 Noise monitoring started...")

    while True:
        if not get_listen_state():
            print("⏸️ Listening paused.")
            time.sleep(1)
            continue

        # --- Core Logic: Detect external changes and adjust ---
        current_system_volume = volume_control.GetMasterVolumeLevelScalar()

        with volume_lock:
            # Check if the volume was changed by an external source (keyboard, other apps)
            # We use a tolerance to account for floating point inaccuracies.
            if abs(current_system_volume - last_volume_set_by_app) > 0.015:
                print(f"[MANUAL DETECTED] System volume changed to {current_system_volume:.2f}. Adopting as new base.")
                # Update the GUI slider to reflect the new reality
                update_base_volume_in_gui(current_system_volume)
                # Store this new value so we don't immediately detect it as another change
                last_volume_set_by_app = current_system_volume

        # --- Auto-adjustment logic ---
        if get_toggle_state(): # Only run if auto-mode is ON
            base_volume = get_base_volume()
            noise = get_noise_level()
            sensitivity = get_sensitivity()

            # Calculate how much to adjust the volume based on noise
            # This can be positive or negative
            noise_delta = noise - RESTING_NOISE_LEVEL
            volume_adjustment = noise_delta * sensitivity

            # Calculate the new target volume by applying the adjustment to the base
            final_volume = base_volume + volume_adjustment
            
            # Clamp the volume to the valid range [0.0, 1.0]
            final_volume = min(MAX_VOLUME, max(MIN_VOLUME, final_volume))

            # Set the new volume
            volume_control.SetMasterVolumeLevelScalar(final_volume, None)
            
            with volume_lock:
                # Record the volume we just set
                last_volume_set_by_app = final_volume
            
            print(f"[AUTO] Base: {base_volume:.2f} | Noise: {noise:.4f} | Final Vol: {final_volume:.2f}")
        
        else: # Auto-mode is OFF
            # When auto is off, ensure our tracker matches the system volume
            with volume_lock:
                last_volume_set_by_app = current_system_volume
            print("[MANUAL MODE] Auto-update is off.")
            time.sleep(0.5)

        time.sleep(UPDATE_INTERVAL)


# --- GUI setup ---
def create_gui(volume_control):
    """Creates and configures the Tkinter GUI."""
    root = tk.Tk()
    root.title("Dynamic Volume Controller")
    root.geometry("420x390")
    root.configure(bg="#1e1e1e")
    root.resizable(False, False)

    # --- Style Configuration ---
    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure("TLabel", background="#1e1e1e", foreground="white", font=("Segoe UI", 12))
    style.configure("TCheckbutton", background="#1e1e1e", foreground="white", font=("Segoe UI", 10))
    style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=8)
    style.configure("Horizontal.TScale", background="#1e1e1e", troughcolor='#555', sliderrelief=tk.FLAT, sliderlength=20)
    
    # --- State variables ---
    sensitivity_var = tk.DoubleVar(value=20.0)
    toggle_var = tk.BooleanVar(value=True)
    listening_var = tk.BooleanVar(value=True)
    # This variable now represents the BASE volume, which the user sets.
    base_volume_var = tk.DoubleVar(value=volume_control.GetMasterVolumeLevelScalar())

    # --- GUI Layout ---
    main_frame = ttk.Frame(root, padding="20 20 20 20")
    main_frame.pack(expand=True, fill="both")
    main_frame.configure(style="TFrame")

    ttk.Label(main_frame, text="🎚️ Fluctuation Sensitivity").pack(pady=(0, 5))
    ttk.Scale(main_frame, from_=1, to=100, orient="horizontal", variable=sensitivity_var, length=300).pack()

    auto_frame = ttk.Frame(main_frame)
    auto_frame.pack(pady=(20, 5))
    ttk.Label(auto_frame, text="🟢 Auto Volume Control").pack(side=tk.LEFT, padx=(0, 10))
    ttk.Checkbutton(auto_frame, text="Enable", variable=toggle_var).pack(side=tk.LEFT)

    ttk.Label(main_frame, text="🎛️ Base Volume Control").pack(pady=(20, 5))
    
    def on_manual_slider_change(value):
        """When the user moves the slider, update the system volume and our tracker."""
        global last_volume_set_by_app
        volume_value = float(value)
        # When auto-mode is off, the slider directly controls system volume.
        if not toggle_var.get():
            volume_control.SetMasterVolumeLevelScalar(volume_value, None)
        # Always update our tracker to reflect the user's intended base.
        with volume_lock:
            last_volume_set_by_app = volume_value
        print(f"[GUI] Base volume set to: {volume_value:.2f}")

    manual_slider = ttk.Scale(
        main_frame,
        from_=MIN_VOLUME,
        to=MAX_VOLUME,
        orient="horizontal",
        variable=base_volume_var, # The slider now controls the base volume
        length=300,
        command=on_manual_slider_change
    )
    manual_slider.pack()

    def toggle_listening():
        current = listening_var.get()
        listening_var.set(not current)
        btn_text.set("⏸️ Pause Listening" if not current else "▶️ Resume Listening")

    btn_text = tk.StringVar(value="▶️ Resume Listening")
    toggle_listening() # Set initial state
    ttk.Button(main_frame, textvariable=btn_text, command=toggle_listening).pack(pady=(30, 10))

    ttk.Label(main_frame, text="Made with ❤️ in Python", font=("Segoe UI", 8)).pack(pady=(15, 5))

    # This function allows the background thread to safely update the GUI's base volume variable.
    def update_base_volume_in_gui(new_volume):
        base_volume_var.set(new_volume)

    return root, sensitivity_var, toggle_var, listening_var, lambda: base_volume_var.get(), update_base_volume_in_gui

# --- Main entry point ---
def main():
    """Initializes and runs the application."""
    global last_volume_set_by_app
    try:
        volume_control = get_volume_control()
        # Initialize the tracker with the current system volume
        with volume_lock:
            last_volume_set_by_app = volume_control.GetMasterVolumeLevelScalar()
    except Exception as e:
        print(f"❌ Error accessing system volume: {e}")
        return

    (root, sensitivity_var, toggle_var, listening_var, 
     get_base_volume_func, update_base_volume_func) = create_gui(volume_control)

    monitor_thread = threading.Thread(
        target=monitor_noise,
        args=(
            volume_control,
            lambda: sensitivity_var.get(),
            lambda: toggle_var.get(),
            lambda: listening_var.get(),
            get_base_volume_func,
            update_base_volume_func
        ),
        daemon=True
    )
    monitor_thread.start()

    root.mainloop()

if __name__ == "__main__":
    main()
