import tkinter as tk
from tkinter import scrolledtext, ttk, Toplevel
from typing import List, Tuple, Any, Dict, Union


class BaseWindow:
    """Utility class for common window operations to keep UI code DRY (Don't Repeat Yourself)."""
    
    @staticmethod
    def set_window_geometry(
        window: Union[tk.Tk, tk.Toplevel], 
        width_pct: float, 
        height_pct: float, 
        min_w: int = 550, 
        min_h: int = 650
    ) -> None:
        """Centers the window on the screen based on screen percentage and enforces minimum sizes."""
        screen_w = window.winfo_screenwidth()
        screen_h = window.winfo_screenheight()
        
        width = max(int(screen_w * width_pct), min_w)
        height = max(int(screen_h * height_pct), min_h)
        
        x = (screen_w // 2) - (width // 2)
        y = (screen_h // 2) - (height // 2)
        
        window.geometry(f"{width}x{height}+{x}+{y}")
        window.minsize(min_w, min_h)


class MainWindow(BaseWindow):
    """Draws the primary input interface (Hub Window) where the user loads files."""
    
    def __init__(self, root: tk.Tk, controller: Any):
        self.root = root
        self.controller = controller
        self.root.title("Fuzzy Rabbit")
        self.set_window_geometry(self.root, width_pct=0.4, height_pct=0.7)
        
        self.excel_input_var = tk.StringVar()
        self.txt_path_var = tk.StringVar()
        self._build_ui()
        
        
    def _build_ui(self):
        """Constructs the main window widgets with consistent padding."""
        # Main container with a standard 20px outer margin
        main_pad = tk.Frame(self.root, padx=20, pady=20)
        main_pad.pack(fill="both", expand=True)
        
        # Excel File Selection
        tk.Label(main_pad, text="Spreadsheet", font=("Arial", 11, "bold")).pack(pady=(0, 5))
        excel_frame = tk.Frame(main_pad)
        excel_frame.pack(fill="x", pady=(0, 15))
        tk.Button(excel_frame, text="Select file.xlsx", command=self.controller.select_excel, fg="#2980b9", font=("Arial", 10, "bold")).pack(side="left", padx=(0, 10))
        tk.Entry(excel_frame, textvariable=self.excel_input_var, font=("Arial", 10)).pack(side="left", fill="x", expand=True, ipady=5)

        # Text File Selection
        tk.Label(main_pad, text="List", font=("Arial", 11, "bold")).pack(pady=(0, 5))
        list_frame = tk.Frame(main_pad)
        list_frame.pack(fill="x", pady=(0, 15))
        tk.Button(list_frame, text="Select file.txt", command=self.controller.select_text, fg="#2980b9", font=("Arial", 10, "bold")).pack(side="left", padx=(0, 10))
        tk.Entry(list_frame, textvariable=self.txt_path_var, font=("Arial", 10)).pack(side="left", fill="x", expand=True, ipady=5)

        # Scrolled Text Area for previewing the text file
        self.input_area = scrolledtext.ScrolledText(main_pad, height=5, font=("Arial", 10))
        self.input_area.pack(fill="both", expand=True, pady=(0, 15))

        # Execution Button
        self.check_button = tk.Button(main_pad, text="Check Similarities", command=self.controller.run_check, fg="#2980b9", font=("Arial", 10, "bold"), height=2)
        self.check_button.pack(pady=(10, 0))

        # Progress Bar (Hidden by default)
        self.progress_bar = ttk.Progressbar(main_pad, orient="horizontal", mode="determinate")
        
    def show_progress(self):
        """Displays the progress bar during processing."""
        self.progress_bar.pack(fill="x", pady=(15, 0))
        
    def hide_progress(self):
        """Hides the progress bar once processing is complete."""
        self.progress_bar.pack_forget()


class ReviewWindow(BaseWindow):
    """Draws the dynamic review interface and captures the user's manual selections."""
    
    def __init__(self, root: tk.Tk, controller: Any, results: List[Any]):
        self.root = root
        self.controller = controller
        self.results = results
        self.selection_map: Dict[str, tk.StringVar] = {}
        
        self.window = Toplevel(self.root)
        self.window.title("Review Matches")
        self.set_window_geometry(self.window, width_pct=0.45, height_pct=0.8)
        
        # Override the native OS 'X' button to ensure clean background threading cleanup
        self.window.protocol("WM_DELETE_WINDOW", self.close)
        
        self._build_ui()

    def _build_ui(self) -> None:
        """Constructs the scrollable canvas and populates it with result groups."""
        # Main container with standard 20px padding
        container = tk.Frame(self.window, padx=20, pady=20)
        container.pack(fill="both", expand=True)
        
        # Setup Scrollable Canvas
        self.canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        scroll_frame = tk.Frame(self.canvas)
        
        canvas_window = self.canvas.create_window((0,0), window=scroll_frame, anchor="nw")
        
        # Bindings to ensure the canvas dynamically resizes with the window
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(canvas_window, width=e.width))
        scroll_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        # Bind MouseWheel scrolling
        self.window.bind_all("<MouseWheel>", self._on_mousewheel)

        items_displayed = False

        # Populate the scrollable frame with match results
        for res in self.results:
            top_score = res.candidates[0].score if res.candidates else 0.0

            # --- HIDDEN HIGH CONFIDENCE LOGIC ---
            # If confidence is > 98%, auto-select the best candidate and hide from UI
            if top_score > 0.98:
                auto_choice = tk.StringVar(value=res.candidates[0].suggested)
                self.selection_map[res.original] = auto_choice
                continue
            # ------------------------------------

            items_displayed = True
            item_container = tk.Frame(scroll_frame, bd=1, relief="groove", pady=10, padx=15)
            item_container.pack(pady=(0, 15), fill="x")
            
            is_low_conf = top_score < 0.85
            initial_val = "NONE" if (is_low_conf or not res.candidates) else res.candidates[0].suggested
            choice_var = tk.StringVar(value=initial_val)
        
            # Color coding based on confidence
            if top_score < 0.50: color = "#c0392b"      # Red (Terrible)
            elif top_score < 0.75: color = "#d35400"    # Orange (Poor)
            elif top_score < 0.92: color = "#d4ac0d"    # Yellow (Okay)
            else: color = "#27ae60"                     # Green (Great)

            tk.Label(item_container, text=f"• {res.original}", font=("Arial", 11, "bold"), fg=color).pack(anchor="center", pady=(0, 10))

            opts_frame = tk.Frame(item_container)
            opts_frame.pack()

            tk.Radiobutton(opts_frame, text="Skip", variable=choice_var, value="NONE", font=("Arial", 10)).pack(anchor="w", pady=2)

            for c in res.candidates:
                c_date = getattr(c, 'date', 'N/A')
                display_label = f"[{int(c.score*100)}%] {c.suggested}  >  {c_date}"
                
                # The 'value' remains purely the text so downstream logic doesn't break
                tk.Radiobutton(opts_frame, text=display_label, variable=choice_var, value=c.suggested, font=("Arial", 10)).pack(anchor="w", pady=2)
            
            # Map the original search item to the Tkinter StringVar containing the choice
            self.selection_map[res.original] = choice_var

        # If everything was >98% and nothing was drawn, display a success message
        if not items_displayed:
            tk.Label(scroll_frame, text="All matches found with >98% confidence!\nClick 'Confirm Selections' to proceed.", font=("Arial", 12, "italic"), fg="#7f8c8d").pack(pady=50)

        # Pack the canvas and scrollbar last so they expand properly
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Footer Actions
        footer = tk.Frame(self.window)
        footer.pack(side="bottom", fill="x", padx=20, pady=(0, 20))
        
        self.progress_bar = ttk.Progressbar(footer, orient="horizontal", mode="determinate")
        self.progress_bar['maximum'] = len(self.results)
        
        self.confirm_btn = tk.Button(footer, text="Confirm Selections", command=self._trigger_confirm, fg='#27ae60', font=("Arial", 10, "bold"), height=2)
        self.confirm_btn.pack(pady=(10, 0))

    def _on_mousewheel(self, event) -> None:
        """Handles mouse wheel scrolling inside the canvas."""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _trigger_confirm(self) -> None:
        """Locks the UI, shows progress, and passes selections back to the controller."""
        self.confirm_btn.config(state="disabled")
        self.progress_bar.pack(fill="x", pady=(10, 0))
        
        # Extract purely the strings from the Tkinter StringVars
        extracted_choices = {orig: var.get() for orig, var in self.selection_map.items()}
        self.controller.process_review_selections(extracted_choices)

    def update_progress(self, value: int) -> None:
        """Called by the background thread to update UI progress safely."""
        self.progress_bar.config(value=value)

    def close(self) -> None:
        """Cleans up bindings to prevent memory leaks before destroying."""
        self.window.unbind_all("<MouseWheel>")
        self.window.destroy()


class ResultsWindow(BaseWindow):
    """Draws the final data table. No business logic allowed here."""
    
    def __init__(self, root: tk.Tk, controller: Any, data: List[Tuple[str, str, str]]):
        self.root = root
        self.controller = controller
        self.data = data
        
        self.window = Toplevel(self.root)
        self.window.title("Fuzzy Rabbit - Results")
        self.set_window_geometry(self.window, width_pct=0.45, height_pct=0.6)
        
        self._build_ui()

    def _build_ui(self) -> None:
        """Constructs the Treeview table for final output."""
        # Top Label
        tk.Label(self.window, text="RESULTS", font=("Arial", 16, "bold")).pack(pady=(20, 5))

        # Table Container with consistent padding
        table_frame = tk.Frame(self.window, padx=20)
        table_frame.pack(fill="both", expand=True)

        columns = ("original", "match", "date")
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")

        tree.heading("original", text="ITEM")
        tree.heading("match", text="MATCH")
        tree.heading("date", text="DATE")

        tree.column("original", anchor="w", width=160, stretch=True)
        tree.column("match", anchor="w", width=160, stretch=True)
        tree.column("date", anchor="center", width=85, stretch=False)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Styling the Treeview
        style = ttk.Style()
        style.configure("Treeview", font=("Arial", 10), rowheight=25)
        tree.tag_configure('oddrow', background='#f7f7f7')
        tree.tag_configure('evenrow', background='white')

        # Insert data
        for i, (orig, match, date) in enumerate(self.data):
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            tree.insert("", "end", values=(orig, match, date), tags=(tag,))

        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Footer Actions with consistent padding
        footer = tk.Frame(self.window, padx=20, pady=20)
        footer.pack(fill="x")

        self.json_btn = tk.Button(
            footer, text="Export JSON", 
            command=self._export_json_action, 
            fg="#2980b9", font=("Arial", 10, "bold"), width=15
        )
        self.json_btn.pack(side="left")

        # Utility Button
        self.copy_btn = tk.Button(
            footer, text="Copy Dates", 
            command=self._copy_dates, 
            fg="#333333", font=("Arial", 10, "bold"), width=15
        )
        self.copy_btn.pack(side="right")
        
    def _export_json_action(self) -> None:
        """Formats the table data into a list of dictionaries for the controller."""
        export_list = []
        for item in self.data:
            # item[1] is the matched name, item[2] is the date
            export_list.append({
                "name": item[1],
                "date": item[2]
            })
        self.controller.save_json_file(export_list)
        self.json_btn.config(text="Done ✓", fg="#27ae60")
        
    def _copy_dates(self) -> None:
        """Extracts only the date column and copies it to the system clipboard."""
        just_dates = "\n".join([item[2] for item in self.data])
        self.window.clipboard_clear()
        self.window.clipboard_append(just_dates)
        self.window.update()
        self.copy_btn.config(text="Done ✓", fg="#27ae60")