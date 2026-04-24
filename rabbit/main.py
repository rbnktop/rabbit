import tkinter as tk
from controller import AppController
from views import SplashScreen

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw() # Hide the main window initially
    splash = SplashScreen(root, "splash.png", duration_ms=2500)
    app = AppController(root)
    root.mainloop()