# effects.py
import time

def window_shake_effect(window):
    x = window.winfo_x()
    y = window.winfo_y()
    for _ in range(5):
        window.geometry(f"+{x + 10}+{y}")
        window.update()
        time.sleep(0.05)
        window.geometry(f"+{x - 10}+{y}")
        window.update()
        time.sleep(0.05)
    window.geometry(f"+{x}+{y}")
