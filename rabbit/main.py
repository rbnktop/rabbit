import tkinter as tk
from controller import AppController

if __name__ == "__main__":
    root = tk.Tk()
    app = AppController(root)
    root.mainloop()