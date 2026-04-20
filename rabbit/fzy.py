"""
Fuzzy Rabbit - Global Responsive Edition (Single Date Optimization)
==================================================================

Main Workflow:
1. User selects Excel file (data source)
2. User selects or pastes text file (search items)
3. Application performs fuzzy matching with similarity scoring
4. User reviews and confirms matches
5. Results displayed in table with option to copy dates
"""

from dataclasses import dataclass
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, Toplevel
from tkinter import ttk
from difflib import SequenceMatcher
import threading

# ======================== Data Models ========================

@dataclass
class MatchCandidate:
    """Represents a single match candidate with its similarity score."""
    suggested: str  # The matched value from Excel pool
    score: float    # Similarity score (0.0 to 1.0)

@dataclass
class MultiMatchResult:
    """Represents search results for a single item with multiple candidate matches."""
    original: str  # The original search term
    candidates: list[MatchCandidate]  # List of top matches sorted by score

# ======================== Utility Functions ========================

def get_similarity(a, b):
    """
    Calculate similarity ratio between two strings using SequenceMatcher.
    Uses case-insensitive comparison for more flexible matching.
    
    Args:
        a: First string to compare
        b: Second string to compare
    
    Returns:
        float: Similarity ratio between 0.0 (no match) and 1.0 (perfect match)
    """
    return SequenceMatcher(None, str(a).lower(), str(b).lower()).ratio()

def load_excel_data(path):
    """
    Load data from an Excel file (.xlsx format).
    
    Args:
        path: File path to the Excel file
    
    Returns:
        pandas.DataFrame: DataFrame containing the Excel data
    
    Raises:
        Exception: If file cannot be read or is not valid Excel format
    """
    return pd.read_excel(path)

def parse_search_items(text):
    """
    Parse multi-line text input into a list of search items.
    Strips whitespace and filters empty lines.
    
    Args:
        text: Multi-line string containing search items (one per line)
    
    Returns:
        list[str]: List of cleaned search items
    """
    return [line.strip() for line in text.splitlines() if line.strip()]

def build_excel_pool(df):
    """
    Extract all unique values from DataFrame to create a searchable pool.
    Converts all values to strings and removes duplicates.
    
    Args:
        df: pandas DataFrame to extract values from
    
    Returns:
        list[str]: List of unique values from all cells in the DataFrame
    """
    return [str(x).strip() for x in pd.unique(df.values.ravel()) if pd.notna(x)]

def find_smart_matches(search_items, excel_pool, threshold=0.4, progress_callback=None):
    """
    Perform fuzzy matching between search items and Excel pool.
    For each search item, finds top 3 matches above the similarity threshold.
    
    Args:
        search_items: List of items to search for
        excel_pool: List of values to search in (from Excel)
        threshold: Minimum similarity score to consider a match (default 0.4)
        progress_callback: Optional callback function(current, total) for progress updates
    
    Returns:
        list[MultiMatchResult]: Results containing candidates for each search item
    """
    results = []
    for i, item in enumerate(search_items):
        # Calculate similarity scores for all candidates above threshold
        all_scores = [MatchCandidate(v, get_similarity(item, v)) for v in excel_pool if get_similarity(item, v) >= threshold]
        all_scores.sort(key=lambda x: x.score, reverse=True)
        # Keep only top 3 candidates
        results.append(MultiMatchResult(original=item, candidates=all_scores[:3]))
        # Call progress callback if provided
        if progress_callback: progress_callback(i + 1, len(search_items))
    return results

def extract_dates_for_match(df, match_value):
    """
    Find the most recent date associated with a matched value in DataFrame.
    Searches all columns (treated as dates) for the match value and returns the maximum date.
    
    Args:
        df: pandas DataFrame where column names are dates
        match_value: The value to search for in the DataFrame
    
    Returns:
        str: ISO format date string (YYYY-MM-DD) of the most recent match, 
             or "No date found" if no match exists
    """
    if not match_value or match_value == "NONE": return "No date found"
    
    match_str = str(match_value).strip()
    found_dates = []
    
    # Search each column for the match value
    for col_name in df.columns:
        if df[col_name].astype(str).str.strip().eq(match_str).any():
            try: 
                found_dates.append(pd.to_datetime(col_name))
            except: 
                continue
    
    # Return only the maximum (most recent) date
    return str(max(found_dates).date()) if found_dates else "No date found"

# ======================== Main Application ========================

class FuzzyFinderApp:
    """
    Main application class managing the GUI and fuzzy matching workflow.
    Handles file selection, similarity matching, user review, and results display.
    """
    
    def __init__(self):
        """Initialize the main application window and instance variables."""
        self.root = tk.Tk()
        self.root.title("Fuzzy Rabbit")
        
        # Set initial window size and position (40% width, 70% height)
        self.set_window_geometry(self.root, width_pct=0.4, height_pct=0.7)
        
        # Variables to store file paths and content
        self.excel_input_var = tk.StringVar()
        self.txt_path_var = tk.StringVar()
        self.input_area = None
        self.processing = False  # Flag to prevent concurrent processing
        
        self.build_main_window()

    # ==================== Window Management ====================

    def set_window_geometry(self, window, width_pct, height_pct, min_w=550, min_h=650):
        """
        Set window size and position to be responsive and centered on screen.
        
        Args:
            window: tk.Tk or Toplevel window to configure
            width_pct: Width as percentage of screen width (0.0 to 1.0)
            height_pct: Height as percentage of screen height (0.0 to 1.0)
            min_w: Minimum width in pixels (default 550)
            min_h: Minimum height in pixels (default 650)
        """
        screen_w = window.winfo_screenwidth()
        screen_h = window.winfo_screenheight()
        width = max(int(screen_w * width_pct), min_w)
        height = max(int(screen_h * height_pct), min_h)
        x = (screen_w // 2) - (width // 2)
        y = (screen_h // 2) - (height // 2)
        window.geometry(f"{width}x{height}+{x}+{y}")
        window.minsize(min_w, min_h)

    # ==================== UI Helper Methods ====================

    def _create_progress_bar(self, parent):
        """
        Create a standardized horizontal progress bar.
        Centralizes progress bar styling for consistency across the application.
        
        Args:
            parent: Parent widget (Frame or Window) to contain the progress bar
        
        Returns:
            ttk.Progressbar: Configured progress bar widget
        """
        return ttk.Progressbar(parent, orient="horizontal", mode="determinate")

    def _show_progress_bar(self, progress_bar, padx=0, pady=10):
        """
        Display a progress bar with consistent padding.
        Centralizes the packing/display logic to avoid repetition.
        
        Args:
            progress_bar: ttk.Progressbar widget to display
            padx: Horizontal padding in pixels (default 0)
            pady: Vertical padding in pixels (default 10)
        """
        progress_bar.pack(fill="x", padx=padx, pady=pady)

    def _hide_progress_bar(self, progress_bar):
        """
        Hide a progress bar from the display.
        
        Args:
            progress_bar: ttk.Progressbar widget to hide
        """
        progress_bar.pack_forget()

    def _create_label(self, parent, text, font_size=10, bold=False, fg="black"):
        """
        Create a standardized label to avoid repeated font specifications.
        Reduces redundancy in label creation across the application.
        
        Args:
            parent: Parent widget to contain the label
            text: Label text content
            font_size: Font size in points (default 10)
            bold: Whether to use bold font (default False)
            fg: Foreground color (default "black")
        
        Returns:
            tk.Label: Configured label widget
        """
        weight = "bold" if bold else "normal"
        return tk.Label(parent, text=text, font=("Arial", font_size, weight), fg=fg)

    def _create_button(self, parent, text, command=None, fg="#2980b9", bg=None, bold=True, height=1):
        """
        Create a standardized button to avoid repeated button specifications.
        Reduces redundancy in button creation across the application.
        
        Args:
            parent: Parent widget to contain the button
            text: Button label text
            command: Callback function when button is clicked
            fg: Foreground (text) color (default blue "#2980b9")
            bg: Background color (default None for system default)
            bold: Whether to use bold font (default True)
            height: Button height in lines (default 1)
        
        Returns:
            tk.Button: Configured button widget
        """
        weight = "bold" if bold else "normal"
        kwargs = {"text": text, "command": command, "fg": fg, "font": ("Arial", 10, weight), "height": height}
        if bg:
            kwargs["bg"] = bg
        return tk.Button(parent, **kwargs)

    # ==================== Main Window ====================

    def build_main_window(self):
        """
        Construct the main application window with file selection and input areas.
        Sets up:
        - Excel file selection
        - Search items input (file or direct text)
        - Check Similarities button
        - Progress bar for matching operations
        """
        main_pad = tk.Frame(self.root, padx=25, pady=20)
        main_pad.pack(fill="both", expand=True)
        # Input section
        # Spreadsheet
        self._create_label(main_pad, "Spreadsheet", font_size=11, bold=True).pack(pady=(0, 5))
        excel_frame = tk.Frame(main_pad)
        excel_frame.pack(fill="x", pady=(0, 15))
        self._create_button(excel_frame, "Select file.xlsx", 
                           command=lambda: self.select_file(self.excel_input_var, ("Excel files", "*.xlsx"))).pack(side="left", padx=(0, 10))
        tk.Entry(excel_frame, textvariable=self.excel_input_var, font=("Arial", 10), width=40).pack(side="left", fill="x", expand=True, ipady=5)

        # List/Search 
        self._create_label(main_pad, "List", font_size=11, bold=True).pack(pady=(0, 5))
        list_frame = tk.Frame(main_pad)
        list_frame.pack(fill="x", pady=(0, 15))
        self._create_button(list_frame, "Select file.txt", 
                           command=lambda: self.select_file(self.txt_path_var, ("Text files", "*.txt"), True)).pack(side="left", padx=(0, 10))
        tk.Entry(list_frame, textvariable=self.txt_path_var, font=("Arial", 10), width=40).pack(side="left", fill="x", expand=True, ipady=5)

        # Text input area for direct paste
        self.input_area = scrolledtext.ScrolledText(main_pad, height=5, width=65)
        self.input_area.pack(fill="y", expand=True, pady=10)
        # End input section
        
        # Primary action button
        self.check_button = self._create_button(main_pad, "Check Similarities", 
                                               command=self.run_similarity_check, 
                                               fg="#2980b9", height=2)
        self.check_button.pack(pady=10)

        # Progress bar (initially hidden)
        self.progress_bar = self._create_progress_bar(main_pad)
        self._show_progress_bar(self.progress_bar)
        self._hide_progress_bar(self.progress_bar)

        # Footer
        self._create_label(main_pad, "Cabelo LTDA", font_size=8, bold=True, fg="gray").pack(side="bottom")

    # ==================== File & Input Handling ====================

    def select_file(self, target_var, file_types, is_txt=False):
        """
        Open file dialog and handle file selection.
        If selecting a text file, automatically populates the input area.
        
        Args:
            target_var: tk.StringVar to store the selected file path
            file_types: Tuple of (description, pattern) for file dialog filter
            is_txt: If True, read file content into input_area (default False)
        """
        path = filedialog.askopenfilename(filetypes=[file_types])
        if path:
            target_var.set(path)
            if is_txt:
                with open(path, "r", encoding="utf-8") as f:
                    self.input_area.delete("1.0", tk.END)
                    self.input_area.insert(tk.END, f.read())

    # ==================== Similarity Checking ====================

    def run_similarity_check(self):
        """
        Validate inputs and initiate the fuzzy matching process in a background thread.
        Prevents concurrent operations by checking the processing flag.
        """
        if self.processing: return
        
        e_path = self.excel_input_var.get()
        t_content = self.input_area.get("1.0", tk.END).strip()
        
        # Validate that both Excel file and search items are provided
        if not e_path or not t_content:
            messagebox.showwarning("Warning", "Input data missing.")
            return
        
        self.processing = True
        self.check_button.config(text="Checking...", state="disabled")
        self._show_progress_bar(self.progress_bar)
        threading.Thread(target=self._exec_check, args=(e_path, t_content), daemon=True).start()

    def _exec_check(self, path, content):
        """
        Execute the fuzzy matching operation in background thread.
        Safely updates UI from thread using root.after() to avoid threading issues.
        
        Args:
            path: Path to Excel file
            content: Multi-line text content with search items
        """
        try:
            df = load_excel_data(path)
            items = parse_search_items(content)
            pool = build_excel_pool(df)
            
            self.progress_bar['maximum'] = len(items)
            
            # Perform fuzzy matching with progress callback
            results = find_smart_matches(items, pool, 
                                        progress_callback=lambda c, t: self.root.after(0, 
                                                                     lambda: self.progress_bar.config(value=c)))
            self.root.after(0, lambda: self._on_check_done(results, df))
        except Exception as e:
            self.root.after(0, lambda: [messagebox.showerror("Error", str(e)), self._reset_ui()])

    def _on_check_done(self, results, df):
        """
        Handle completion of matching operation.
        Sorts results by confidence and opens the review window.
        
        Args:
            results: List of MultiMatchResult objects
            df: Original DataFrame for date extraction
        """
        self._reset_ui()
        # Sort by match confidence (lowest scores first for user review)
        results.sort(key=lambda x: x.candidates[0].score if x.candidates else 0)
        self.open_review_window(results, df)

    def _reset_ui(self):
        """
        Reset UI elements after matching is complete.
        Re-enables the check button and hides the progress bar.
        """
        self.processing = False
        self.check_button.config(state="normal")
        self._hide_progress_bar(self.progress_bar)

    # ==================== Review Window ====================

    def open_review_window(self, results, df):
        """
        Create interactive window for user to review and confirm/adjust matches.
        Displays each search item with its top 3 candidate matches.
        User can select alternatives or skip items.
        
        Args:
            results: List of MultiMatchResult objects with candidates
            df: DataFrame for later date extraction
        """
        rw = Toplevel(self.root)
        rw.title("Review Matches")
        rw.withdraw()  # Hide until fully constructed
        
        self.set_window_geometry(rw, width_pct=0.6, height_pct=0.8, min_w=700, min_h=500)

        # Cleanup handler
        def on_close():
            rw.unbind_all("<MouseWheel>")
            rw.destroy()
        rw.protocol("WM_DELETE_WINDOW", on_close)
        
        # Main container
        container = tk.Frame(rw)
        container.pack(fill="both", expand=True, padx=10, pady=10)
        self._create_label(container, "Review Items", font_size=11, bold=True).pack(pady=(0, 5))
        
        # Scrollable area setup
        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas)
        
        canvas_window = canvas.create_window((0,0), window=scroll_frame, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_window, width=e.width))
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.configure(yscrollcommand=scrollbar.set)
        rw.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        # Build match items list
        selection_map = []  # Stores (original_item, tk.StringVar) pairs
        for res in results:
            item_container = tk.Frame(scroll_frame, bd=1, relief="groove", pady=5)
            item_container.pack(fill="x", pady=2, padx=5)
            
            header = tk.Frame(item_container)
            header.pack(fill="x", padx=5)

            # Initialize with best match or "NONE" if no candidates
            choice_var = tk.StringVar(value=res.candidates[0].suggested if res.candidates else "NONE")
            
            # Display original search item
            self._create_label(header, f"• {res.original}", font_size=10, bold=True).pack(side="left")
            
            # Display current selection (updates when changed)
            curr_lbl = tk.Label(header, text=f"→ {choice_var.get()}", fg="#2980b9", font=("Arial", 9, "italic"))
            curr_lbl.pack(side="left", padx=10)

            # Collapsible alternatives panel
            details = tk.Frame(item_container, bg="#f9f9f9", pady=5)
            btn = tk.Button(header, text="▼ Alternatives", font=("Arial", 8), relief="flat", fg="gray", 
                            command=lambda d=details: d.pack(fill="x", padx=20) if not d.winfo_viewable() else d.pack_forget())
            btn.pack(side="right")

            # Build radio button options
            if res.candidates:
                # Skip option
                tk.Radiobutton(details, text="Skip", variable=choice_var, value="NONE", bg="#f9f9f9", 
                              command=lambda v=choice_var, l=curr_lbl: l.config(text=f"→ {v.get()}")).pack(anchor="w")
                # Match options with confidence scores
                for c in res.candidates:
                    tk.Radiobutton(details, text=f"[{int(c.score*100)}%] {c.suggested}", 
                                  variable=choice_var, value=c.suggested, bg="#f9f9f9", 
                                  command=lambda v=choice_var, l=curr_lbl: l.config(text=f"→ {v.get()}")).pack(anchor="w")
            else: 
                tk.Label(details, text="No matches found.", fg="red", bg="#f9f9f9").pack(anchor="w")
            
            selection_map.append((res.original, choice_var))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Footer with confirm button and progress bar
        footer = tk.Frame(rw, height=80)
        footer.pack(side="bottom")
        progress_bar = self._create_progress_bar(footer)
        
        def run_confirm_bt():
            """Start final processing of confirmed selections."""
            confirm_btn.config(state="disabled")
            self._show_progress_bar(progress_bar, pady=10)
            confirm_btn.config(text="Confirming...")
            threading.Thread(target=self._final_task, args=(selection_map, df, rw, progress_bar), daemon=True).start()

        confirm_btn = self._create_button(footer, "Confirm Selections", command=run_confirm_bt, fg='#50906b', height=2)
        confirm_btn.pack(padx=20, pady=10)
        
        rw.deiconify()  # Show window now that construction is complete

    # ==================== Final Processing ====================

    def _final_task(self, selection_map, df, win, progress_bar):
        """
        Extract dates for confirmed selections and prepare results for display.
        Runs in background thread to keep UI responsive.
        
        Args:
            selection_map: List of (original_item, tk.StringVar) tuples with user selections
            df: DataFrame for date extraction
            win: Review window to close after processing
            progress_bar: Progress bar to update during processing
        """
        progress_bar['maximum'] = len(selection_map)
        final_results = []
        
        for i, (orig, var) in enumerate(selection_map):
            choice = var.get()
            match_display = choice if choice != "NONE" else "---"
            # Extract date for the selected match (or "No date found" if skipped)
            date = extract_dates_for_match(df, choice) if choice != "NONE" else "No date found"
            
            final_results.append((orig, match_display, date))
            self.root.after(0, lambda v=i+1: progress_bar.config(value=v))
        
        # Update UI safely from thread
        self.root.after(0, lambda: win.unbind_all("<MouseWheel>"))
        self.root.after(0, lambda: [self.show_results(final_results), win.destroy()])

    # ==================== Results Display ====================

    def show_results(self, data):
        """
        Display final results in a formatted table with copy-to-clipboard functionality.
        
        Args:
            data: List of (original, match, date) tuples to display
        """
        res_win = Toplevel(self.root)
        res_win.title("Fuzzy Rabbit - Results")
        self.set_window_geometry(res_win, width_pct=0.75, height_pct=0.7)

        self._create_label(res_win, "RESULTS", font_size=16, bold=True).pack(pady=10)

        # Table container
        table_frame = tk.Frame(res_win)
        table_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # Treeview table setup
        columns = ("original", "match", "date")
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")

        tree.heading("original", text="ITEM")
        tree.heading("match", text="MATCH")
        tree.heading("date", text="DATE")

        tree.column("original", anchor="w", width=250)
        tree.column("match", anchor="w", width=250)
        tree.column("date", anchor="center", width=200)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Style table rows (alternating colors for readability)
        style = ttk.Style()
        style.configure("Treeview", font=("Arial", 10), rowheight=28)
        tree.tag_configure('oddrow', background='#f7f7f7')
        tree.tag_configure('evenrow', background='white')

        # Populate table with results
        for i, (orig, match, date) in enumerate(data):
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            tree.insert("", "end", values=(orig, match, date), tags=(tag,))

        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Footer with copy button
        footer = tk.Frame(res_win, pady=20)
        footer.pack(fill="x")

        def copy_dates_bt():
            """Copy only the date column to clipboard and show feedback on button."""
            just_dates = "\n".join([item[2] for item in data])
            res_win.clipboard_clear()
            res_win.clipboard_append(just_dates)
            res_win.update()
            copy_btn.config(text="Done ✓", fg="#27ae60")

        copy_btn = self._create_button(footer, "Copy Dates", command=copy_dates_bt, 
                                      bg="#ecf0f1", bold=True)
        copy_btn.pack()

# ======================== Entry Point ========================

if __name__ == "__main__":
    """Initialize and run the application."""
    FuzzyFinderApp().root.mainloop()

