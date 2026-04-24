import tkinter as tk
from controller import AppController
from views import SplashScreen
import ctypes

# This tells Windows to treat this script as its own application
# allowing the custom icon to show on the taskbar.
myappid = 'CabeloLtda.fuzzyrabbit.vF' # arbitrary string
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw() # Hide the main window initially
    splash = SplashScreen(root, "splash.png", duration_ms=2500)
    app = AppController(root)
    root.mainloop()