import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pynput import keyboard
import time
import threading
import json
import sys

# Prevent command prompt from showing
if sys.platform.startswith('win'):
    import ctypes
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

# Global Variables
recording = False
recorded_keys = []
HOTKEYS = {
    '<ctrl>+r': 'start_recording',
    '<ctrl>+s': 'stop_recording',
    '<ctrl>+p': 'playback',
    '<ctrl>+<shift>+s': 'save_macro',
    '<ctrl>+o': 'load_macro',
    '<ctrl>+l': 'toggle_loop'
}

class MacroRecorder:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Macro Recorder")
        self.root.geometry("400x550")
        self.root.resizable(False, False)
        
        # Additional variables for loop control
        self.loop_enabled = tk.BooleanVar(value=False)
        self.is_playing = False
        self.stop_playback = False
        
        # Apply theme
        style = ttk.Style()
        style.theme_use('clam')
        
        self.setup_ui()
        self.setup_hotkeys()
        self.keyboard_controller = keyboard.Controller()

    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = ttk.Label(
            main_frame, 
            text="Macro Recorder", 
            font=("Helvetica", 16, "bold")
        )
        title_label.pack(pady=10)

        # Status frame
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding="5")
        status_frame.pack(fill=tk.X, pady=10)
        
        self.status_label = ttk.Label(
            status_frame, 
            text="Status: Idle", 
            font=("Helvetica", 10)
        )
        self.status_label.pack(pady=5)

        # Buttons frame
        buttons_frame = ttk.LabelFrame(main_frame, text="Controls", padding="10")
        buttons_frame.pack(fill=tk.BOTH, expand=True)

        # Button configurations
        button_configs = [
            ("Start Recording", "start_recording", "Ctrl+R"),
            ("Stop Recording", "stop_recording", "Ctrl+S"),
            ("Playback", "playback", "Ctrl+P"),
            ("Save Macro", "save_macro", "Ctrl+Shift+S"),
            ("Load Macro", "load_macro", "Ctrl+O")
        ]

        for text, command, shortcut in button_configs:
            frame = ttk.Frame(buttons_frame)
            frame.pack(fill=tk.X, pady=5)
            
            btn = ttk.Button(
                frame,
                text=text,
                command=getattr(self, command)
            )
            btn.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            shortcut_label = ttk.Label(
                frame,
                text=shortcut,
                font=("Helvetica", 8),
                foreground="gray"
            )
            shortcut_label.pack(side=tk.RIGHT, padx=5)

        # Playback controls frame
        playback_frame = ttk.LabelFrame(main_frame, text="Playback Settings", padding="5")
        playback_frame.pack(fill=tk.X, pady=10)

        # Loop checkbox with keyboard shortcut hint
        loop_frame = ttk.Frame(playback_frame)
        loop_frame.pack(fill=tk.X, pady=5)
        
        self.loop_checkbox = ttk.Checkbutton(
            loop_frame,
            text="Loop Playback",
            variable=self.loop_enabled,
            command=self.toggle_loop
        )
        self.loop_checkbox.pack(side=tk.LEFT)
        
        loop_shortcut = ttk.Label(
            loop_frame,
            text="Ctrl+L",
            font=("Helvetica", 8),
            foreground="gray"
        )
        loop_shortcut.pack(side=tk.RIGHT, padx=5)

        # Stop button for loop playback
        self.stop_button = ttk.Button(
            playback_frame,
            text="Stop Playback",
            command=self.stop_loop_playback,
            state=tk.DISABLED
        )
        self.stop_button.pack(fill=tk.X, pady=5)

        # Info frame
        info_frame = ttk.LabelFrame(main_frame, text="Information", padding="5")
        info_frame.pack(fill=tk.X, pady=10)
        
        self.info_label = ttk.Label(
            info_frame,
            text="Keys recorded: 0",
            font=("Helvetica", 10)
        )
        self.info_label.pack(pady=5)

    def setup_hotkeys(self):
        def for_canonical(f):
            return lambda k: f(self.listener.canonical(k))

        hotkey_handler = keyboard.GlobalHotKeys({
            '<ctrl>+r': self.start_recording,
            '<ctrl>+s': self.stop_recording,
            '<ctrl>+p': self.playback,
            '<ctrl>+<shift>+s': self.save_macro,
            '<ctrl>+o': self.load_macro,
            '<ctrl>+l': self.toggle_loop
        })
        hotkey_handler.start()

    def toggle_loop(self):
        # Toggle loop mode using keyboard shortcut or checkbox
        self.loop_enabled.set(not self.loop_enabled.get())

    def stop_loop_playback(self):
        self.stop_playback = True

    def start_recording(self):
        global recording, recorded_keys
        if recording:
            messagebox.showinfo("Info", "Already recording!")
            return
        recorded_keys.clear()
        recording = True
        threading.Thread(target=self.record_keys, daemon=True).start()
        self.status_label.config(text="Status: Recording...")
        self.info_label.config(text="Keys recorded: 0")

    def stop_recording(self):
        global recording
        if not recording:
            messagebox.showinfo("Info", "Not currently recording!")
            return
        recording = False
        self.status_label.config(text="Status: Idle")
        self.info_label.config(text=f"Keys recorded: {len(recorded_keys)}")

    def record_keys(self):
        with keyboard.Listener(on_press=self.on_press, on_release=self.on_release) as self.listener:
            self.listener.join()

    def on_press(self, key):
        global recording, recorded_keys
        if recording:
            try:
                recorded_keys.append((time.time(), 'press', key.char))
            except AttributeError:
                recorded_keys.append((time.time(), 'press', str(key)))
            self.info_label.config(text=f"Keys recorded: {len(recorded_keys)}")

    def on_release(self, key):
        global recording, recorded_keys
        if recording:
            try:
                recorded_keys.append((time.time(), 'release', key.char))
            except AttributeError:
                recorded_keys.append((time.time(), 'release', str(key)))
            self.info_label.config(text=f"Keys recorded: {len(recorded_keys)}")

    def playback(self):
        if not recorded_keys:
            messagebox.showinfo("Info", "No keys recorded!")
            return
        
        if self.is_playing:
            messagebox.showinfo("Info", "Playback already in progress!")
            return
            
        self.is_playing = True
        self.stop_playback = False
        self.status_label.config(text="Status: Playing...")
        self.stop_button.config(state=tk.NORMAL)
        threading.Thread(target=self._playback_thread, daemon=True).start()

    def _playback_thread(self):
        try:
            while not self.stop_playback:
                base_time = recorded_keys[0][0]
                for event in recorded_keys:
                    if self.stop_playback:
                        break
                        
                    event_time, action, key = event
                    time.sleep(event_time - base_time)
                    base_time = event_time
                    
                    # Handle special keys
                    if isinstance(key, str) and key.startswith("Key."):
                        key = keyboard.Key[key.split('.')[1]]
                    
                    if action == 'press':
                        self.keyboard_controller.press(key)
                    elif action == 'release':
                        self.keyboard_controller.release(key)
                
                if not self.loop_enabled.get():
                    break
                    
                # Add a small delay between loops
                time.sleep(0.5)
                
        finally:
            self.is_playing = False
            self.stop_playback = False
            self.root.after(0, lambda: self.status_label.config(text="Status: Idle"))
            self.root.after(0, lambda: self.stop_button.config(state=tk.DISABLED))

    def save_macro(self):
        if not recorded_keys:
            messagebox.showinfo("Info", "No keys recorded to save!")
            return
        filepath = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")]
        )
        if filepath:
            with open(filepath, 'w') as file:
                json.dump(recorded_keys, file)
            messagebox.showinfo("Success", f"Macro saved to {filepath}")

    def load_macro(self):
        global recorded_keys
        filepath = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if filepath:
            with open(filepath, 'r') as file:
                recorded_keys = json.load(file)
            self.info_label.config(text=f"Keys recorded: {len(recorded_keys)}")
            messagebox.showinfo("Success", f"Macro loaded from {filepath}")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = MacroRecorder()
    app.run()