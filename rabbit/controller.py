import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import json
import engine as engine
from views import MainWindow, ReviewWindow, ResultsWindow



class AppController:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.ui = MainWindow(root, self)
        self.processing = False
        self.current_df = None
        self.original_results = []
        self.review_window = None

    def select_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.ui.excel_input_var.set(path)

    def select_text(self):
        path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if path:
            self.ui.txt_path_var.set(path)
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read()
                    self.ui.input_area.delete("1.0", tk.END)
                    self.ui.input_area.insert(tk.END, content)
            except Exception as e:
                messagebox.showerror("File Error", f"Could not read text file: {e}")

    def run_check(self):
        if self.processing: 
            return
        
        e_path = self.ui.excel_input_var.get()
        t_content = self.ui.input_area.get("1.0", tk.END).strip()
        
        if not e_path or not t_content:
            messagebox.showwarning("Warning", "Input data missing.")
            return
            
        self.processing = True
        self.ui.check_button.config(text="Checking...", state="disabled")
        self.ui.show_progress()
        
        threading.Thread(target=self._exec_check, args=(e_path, t_content), daemon=True).start()
    
    def process_review_selections(self, selections: dict):
        """Starts the final date extraction in a background thread."""
        threading.Thread(target=self._compile_final_data, args=(selections,), daemon=True).start()
    
    def save_json_file(self, data: list):
        """Saves the acquisition data as a JSON dictionary."""
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            title="Send this to Swift Hawk"
        )
        if not path:
            return

        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("Success", "File saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save file: {e}")

    def _exec_check(self, path: str, content: str):
        try:
            self.current_df = engine.load_excel_data(path)
            pool = engine.build_excel_pool(self.current_df)
            items = engine.parse_search_items(content)

            self.root.after(0, lambda: self.ui.progress_bar.config(maximum=len(items)))
            
            results = engine.find_smart_matches(
                items, pool, 
                progress_callback=lambda c, t: self.root.after(0, self._update_main_progress, c)
            )
            
            self.root.after(0, self._on_check_done, results)
            
        except Exception as e:
            self.root.after(0, lambda err=e : self._handle_error(err))

    def _handle_error(self, err: Exception):
        messagebox.showerror("Error", str(err))
        self._reset_ui()

    def _on_check_done(self, results):
        self._reset_ui()
        self.original_results = results 
        self.review_window = ReviewWindow(self.root, self, results)

    def _compile_final_data(self, selections: dict):
        final_data = []
        
        for i, res in enumerate(self.original_results):
            choice = selections.get(res.original, "NONE")
            
            # Logic requirement: items in results must be in original input order
            if choice != "NONE" and self.current_df is not None:
                match_display = choice
                date_display = engine.extract_dates_for_match(self.current_df, choice)
            else:
                match_display = "====NaN===="
                date_display = "FOTO"
            
            final_data.append((res.original, match_display, date_display))
            
            # Update the ReviewWindow progress bar if it exists
            if self.review_window is not None:
                self.root.after(0, lambda v=i+1: self._safe_progress_update(v))
            
        self.root.after(0, self._show_results, final_data)

    def _show_results(self, final_data: list):
        if self.review_window:
            self.review_window.close()
        self.results_window = ResultsWindow(self.root, self, final_data)

    def _safe_progress_update(self, value: int):
        """Prevents crashes if the review window is closed mid-process."""
        if self.review_window and hasattr(self.review_window, 'update_progress'):
            self.review_window.update_progress(value)

    def _reset_ui(self):
        self.processing = False
        self.ui.check_button.config(text="Check Similarities", state="normal")
        self.ui.hide_progress()

    def _update_main_progress(self, current: int):
        """Type-safe helper for the main progress bar."""
        self.ui.progress_bar.config(value=current)