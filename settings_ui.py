import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import threading
from logger import get_logger

log = get_logger()

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')


def load_config():
    try:
        with open(CONFIG_PATH, 'r') as f:
            return json.load(f)
    except:
        return {}


def save_config(data):
    try:
        existing = load_config()
        existing.update(data)
        with open(CONFIG_PATH, 'w') as f:
            json.dump(existing, f, indent=2)
        log.info("Settings saved to config.json")
        return True
    except Exception as e:
        log.error("Failed to save settings: %s", e)
        return False


def test_hub_connection(hub_url, agent_key):
    """Test if the hub is reachable and the agent key is valid."""
    import requests
    try:
        headers = {'Authorization': f'Bearer {agent_key}'}
        resp = requests.get(f'{hub_url}/api/print-hub/profiles', headers=headers, timeout=5)
        if resp.status_code == 200:
            profiles = resp.json().get('profiles', {})
            return True, f"Connected! {len(profiles)} profile(s) synced."
        elif resp.status_code == 401:
            return False, "Authentication failed - invalid Agent Key."
        else:
            return False, f"Hub returned status {resp.status_code}"
    except requests.ConnectionError:
        return False, "Cannot reach the Hub server. Check the URL."
    except Exception as e:
        return False, str(e)


def open_settings_window():
    """Opens a tkinter settings dialog from the tray menu."""

    def _run():
        config = load_config()

        root = tk.Tk()
        root.title("Trayprint - Settings")
        root.geometry("520x480")
        root.resizable(False, False)
        root.configure(bg='#1a1d27')

        # Style
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TLabel', background='#1a1d27', foreground='#e4e6ed', font=('Segoe UI', 10))
        style.configure('Header.TLabel', background='#1a1d27', foreground='#818cf8', font=('Segoe UI', 14, 'bold'))
        style.configure('Sub.TLabel', background='#1a1d27', foreground='#8b8fa3', font=('Segoe UI', 9))
        style.configure('Status.TLabel', background='#1a1d27', foreground='#8b8fa3', font=('Segoe UI', 9))
        style.configure('TEntry', fieldbackground='#0f1117', foreground='#e4e6ed', insertcolor='#e4e6ed')
        style.configure('TButton', background='#6366f1', foreground='white', font=('Segoe UI', 10, 'bold'), padding=8)
        style.map('TButton', background=[('active', '#818cf8')])
        style.configure('Test.TButton', background='#22263a', foreground='#e4e6ed', font=('Segoe UI', 9), padding=5)
        style.map('Test.TButton', background=[('active', '#2a2e3f')])

        # Header
        ttk.Label(root, text="Trayprint Settings", style='Header.TLabel').pack(pady=(20, 5))
        ttk.Label(root, text="Connect this agent to your central Print Hub", style='Sub.TLabel').pack(pady=(0, 20))

        # Frame
        frame = tk.Frame(root, bg='#22263a', padx=20, pady=20, highlightbackground='#2a2e3f', highlightthickness=1)
        frame.pack(padx=20, fill='x')

        # Port
        ttk.Label(frame, text="Local Port").pack(anchor='w', pady=(0, 2))
        port_var = tk.StringVar(value=str(config.get('port', 49211)))
        port_entry = ttk.Entry(frame, textvariable=port_var, width=50)
        port_entry.pack(fill='x', pady=(0, 12))

        # Hub URL
        ttk.Label(frame, text="Print Hub URL").pack(anchor='w', pady=(0, 2))
        hub_var = tk.StringVar(value=config.get('hub_url', ''))
        hub_entry = ttk.Entry(frame, textvariable=hub_var, width=50)
        hub_entry.pack(fill='x', pady=(0, 4))
        ttk.Label(frame, text="e.g. http://192.168.1.100:8082", style='Sub.TLabel').pack(anchor='w', pady=(0, 12))

        # Agent Key
        ttk.Label(frame, text="Agent Key").pack(anchor='w', pady=(0, 2))
        key_var = tk.StringVar(value=config.get('agent_key', ''))
        key_entry = ttk.Entry(frame, textvariable=key_var, width=50)
        key_entry.pack(fill='x', pady=(0, 4))
        ttk.Label(frame, text="Copy this from Print Hub → Agents page", style='Sub.TLabel').pack(anchor='w', pady=(0, 12))

        # Status label
        status_var = tk.StringVar(value="")
        status_label = ttk.Label(frame, textvariable=status_var, style='Status.TLabel', wraplength=440)
        status_label.pack(fill='x', pady=(0, 5))

        # Test Connection button
        def on_test():
            hub = hub_var.get().strip().rstrip('/')
            key = key_var.get().strip()
            if not hub or not key:
                status_var.set("Please enter both Hub URL and Agent Key.")
                status_label.configure(foreground='#f59e0b')
                return
            status_var.set("Testing connection...")
            status_label.configure(foreground='#8b8fa3')
            root.update()

            success, msg = test_hub_connection(hub, key)
            status_var.set(msg)
            status_label.configure(foreground='#22c55e' if success else '#ef4444')

        ttk.Button(frame, text="Test Connection", command=on_test, style='Test.TButton').pack(fill='x', pady=(0, 5))

        # Buttons frame
        btn_frame = tk.Frame(root, bg='#1a1d27')
        btn_frame.pack(pady=20, padx=20, fill='x')

        def on_save():
            new_config = {
                'port': int(port_var.get().strip() or 49211),
                'hub_url': hub_var.get().strip().rstrip('/'),
                'agent_key': key_var.get().strip(),
            }
            if save_config(new_config):
                status_var.set("Settings saved! Restart Trayprint to apply changes.")
                status_label.configure(foreground='#22c55e')
            else:
                status_var.set("Failed to save settings.")
                status_label.configure(foreground='#ef4444')

        ttk.Button(btn_frame, text="Save Settings", command=on_save).pack(side='left', expand=True, fill='x', padx=(0, 5))

        cancel_btn = ttk.Button(btn_frame, text="Close", command=root.destroy, style='Test.TButton')
        cancel_btn.pack(side='right', expand=True, fill='x', padx=(5, 0))

        root.mainloop()

    # Run in a separate thread so pystray doesn't block
    threading.Thread(target=_run, daemon=True).start()
