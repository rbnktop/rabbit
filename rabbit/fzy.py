"""
Fuzzy Rabbit - Global Responsive Edition
========================================

A desktop application designed to perform fuzzy matching between a list
of search items and an Excel data source. 

Main Workflow:
1. User selects an Excel file (data source).
2. User selects or pastes a text file (search items).
3. Application performs fuzzy matching with similarity scoring.
4. User reviews and confirms matches.
5. Results are displayed in a table with an option to copy the matched dates.
"""

from dataclasses import dataclass
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, Toplevel
from tkinter import ttk
from difflib import SequenceMatcher
import threading
from typing import List, Tuple, Optional, Callable, Any

# ======================== Data Models ========================

@dataclass
class MatchCandidate:
    """
    Represents a single match candidate with its similarity score.
    
    Attributes:
        suggested (str): The matched value found from the Excel pool.
        score (float): The similarity score ranging from 0.0 to 1.0.
    """
    suggested: str
    score: float

@dataclass
class MultiMatchResult:
    """
    Represents search results for a single item with multiple candidate matches.
    
    Attributes:
        original (str): The original search term provided by the user.
        candidates (List[MatchCandidate]): A list of top match candidates sorted by score.
    """
    original: str
    candidates: List[MatchCandidate]

# ======================== Utility Functions ========================

def get_similarity(a: str, b: str) -> float:
    """
    Calculate the similarity ratio between two strings using SequenceMatcher.
    
    Uses a case-insensitive comparison to provide more flexible matching results.
    
    Args:
        a (str): The first string to compare.
        b (str): The second string to compare.
    
    Returns:
        float: Similarity ratio between 0.0 (no match) and 1.0 (perfect match).
    """
    return SequenceMatcher(None, str(a).lower(), str(b).lower()).ratio()

def load_excel_data(path: str) -> pd.DataFrame:
    """
    Load data from an Excel file (.xlsx format) into a pandas DataFrame.
    
    Args:
        path (str): The absolute or relative file path to the Excel file.
    
    Returns:
        pd.DataFrame: A DataFrame containing the loaded Excel data.
    
    Raises:
        Exception: If the file cannot be read or is not a valid Excel format.
    """
    return pd.read_excel(path)

def parse_search_items(text: str) -> List[str]:
    """
    Parse a multi-line text input string into a list of individual search items.
    
    Strips leading/trailing whitespace and automatically filters out empty lines.
    
    Args:
        text (str): Multi-line string containing search items (one per line).
    
    Returns:
        List[str]: A list of cleaned, non-empty search item strings.
    """
    return [line.strip() for line in text.splitlines() if line.strip()]

def build_excel_pool(df: pd.DataFrame) -> List[str]:
    """
    Extract all unique values from a DataFrame to create a unified searchable pool.
    
    Converts all valid values to strings, strips whitespace, and removes duplicates.
    
    Args:
        df (pd.DataFrame): The pandas DataFrame to extract values from.
    
    Returns:
        List[str]: A list containing all unique string values found in the DataFrame.
    """
    return [str(x).strip() for x in pd.unique(df.values.ravel()) if pd.notna(x)]

def find_smart_matches(
    search_items: List[str], 
    excel_pool: List[str], 
    threshold: float = 0.4, 
    progress_callback: Optional[Callable[[int, int], None]] = None
) -> List[MultiMatchResult]:
    """
    Perform fuzzy matching between a list of search items and the Excel pool.
    
    For each search item, calculates the similarity score against all pool items.
    Retains up to the top 3 matches that meet or exceed the specified threshold.
    
    Args:
        search_items (List[str]): The items the user wants to search for.
        excel_pool (List[str]): The unique values extracted from the Excel file.
        threshold (float): Minimum similarity score (0.0 - 1.0) required to register a match.
        progress_callback (Callable[[int, int], None], optional): A function to report 
            progress, taking current iteration and total iterations as arguments.
            
    Returns:
        List[MultiMatchResult]: A list of results mapping each original search term
            to its highest-scoring candidate matches.
    """
    results = []
    for i, item in enumerate(search_items):
        # Calculate similarity scores for all candidates above the threshold
        all_scores = [
            MatchCandidate(v, get_similarity(item, v)) 
            for v in excel_pool if get_similarity(item, v) >= threshold
        ]
        
        # Sort scores in descending order and keep the top 3
        all_scores.sort(key=lambda x: x.score, reverse=True)
        results.append(MultiMatchResult(original=item, candidates=all_scores[:3]))
        
        if progress_callback:
            progress_callback(i + 1, len(search_items))
            
    return results

def extract_dates_for_match(df: pd.DataFrame, match_value: str) -> str:
    """
    Find the most recent date associated with a matched value in the DataFrame.
    
    Treats DataFrame column headers as dates. Searches the entire DataFrame for 
    the exact matched string, collects all corresponding column header dates, 
    and returns the most recent one formatted as DD/MM/YYYY.
    
    Args:
        df (pd.DataFrame): The DataFrame where column headers represent dates.
        match_value (str): The finalized value chosen by the user to search for.
    
    Returns:
        str: Date string formatted as "DD/MM/YYYY" representing the most recent match, 
             or "No date found" if no occurrences exist.
    """
    if not match_value or match_value == "NONE":
        return "No date found"
    
    match_str = str(match_value).strip()
    found_dates = []
    
    # Iterate through columns to locate instances of the matched value
    for col_name in df.columns:
        if df[col_name].astype(str).str.strip().eq(match_str).any():
            try: 
                found_dates.append(pd.to_datetime(col_name))
            except (ValueError, TypeError): 
                # Ignore column headers that cannot be parsed as a datetime
                continue
    
    # Return the maximum (latest) date formatted as Day/Month/Year
    if found_dates:
        return max(found_dates).strftime("%d/%m/%Y")
    
    return "No date found"

# ======================== Main Application ========================

class FuzzyFinderApp:
    """
    Main application class managing the GUI and fuzzy matching workflow.
    
    Handles file selection, coordinates similarity matching via background threads,
    facilitates user review of candidates, and presents final result data.
    """
    
    def __init__(self) -> None:
        """Initialize the main application window and set up necessary variables."""
        self.root = tk.Tk()
        self.root.title("Fuzzy Rabbit")
        
        # Set initial window size dynamically based on screen resolution
        self.set_window_geometry(self.root, width_pct=0.4, height_pct=0.7)
        
        self.excel_input_var = tk.StringVar()
        self.txt_path_var = tk.StringVar()
        self.input_area: Optional[scrolledtext.ScrolledText] = None
        
        # Flag to prevent multiple simultaneous matching processes
        self.processing = False  
        
        self.build_main_window()

    # ==================== Window Management ====================

    def set_window_geometry(self, window: tk.Wm, width_pct: float, height_pct: float, min_w: int = 550, min_h: int = 650) -> None:
        """
        Set window size and position it dynamically in the center of the screen.
        
        Args:
            window (tk.Wm): The tkinter root or Toplevel window to configure.
            width_pct (float): Desired width as a percentage of screen width (e.g., 0.4).
            height_pct (float): Desired height as a percentage of screen height (e.g., 0.7).
            min_w (int): Absolute minimum width in pixels to prevent over-shrinking.
            min_h (int): Absolute minimum height in pixels to prevent over-shrinking.
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

    def _create_progress_bar(self, parent: Any) -> ttk.Progressbar:
        """
        Create a standardized horizontal progress bar widget.
        
        Args:
            parent (Any): The parent tkinter widget containing the progress bar.
        
        Returns:
            ttk.Progressbar: The configured progress bar instance.
        """
        return ttk.Progressbar(parent, orient="horizontal", mode="determinate")

    def _show_progress_bar(self, progress_bar: ttk.Progressbar, padx: int = 0, pady: int = 10) -> None:
        """
        Display a progress bar within the UI.
        
        Args:
            progress_bar (ttk.Progressbar): The progress bar widget to display.
            padx (int): Horizontal padding.
            pady (int): Vertical padding.
        """
        progress_bar.pack(fill="x", padx=padx, pady=pady)

    def _hide_progress_bar(self, progress_bar: ttk.Progressbar) -> None:
        """
        Remove a progress bar from the visible UI layout.
        
        Args:
            progress_bar (ttk.Progressbar): The progress bar widget to hide.
        """
        progress_bar.pack_forget()

    def _create_label(self, parent: Any, text: str, font_size: int = 10, bold: bool = False, fg: str = "black") -> tk.Label:
        """
        Create a standardized text label.
        
        Args:
            parent (Any): The parent tkinter widget.
            text (str): The string to display in the label.
            font_size (int): The Arial font size to apply.
            bold (bool): Whether the text should be bold.
            fg (str): The hexadecimal or named text color.
        
        Returns:
            tk.Label: The configured label widget.
        """
        weight = "bold" if bold else "normal"
        return tk.Label(parent, text=text, font=("Arial", font_size, weight), fg=fg)

    def _create_button(self, parent: Any, text: str, command: Callable[[], None] = None, fg: str = "#2980b9", bg: Optional[str] = None, bold: bool = True, height: int = 1) -> tk.Button:
        """
        Create a standardized interactive button.
        
        Args:
            parent (Any): The parent tkinter widget.
            text (str): The text to display on the button.
            command (Callable): The function to execute upon button click.
            fg (str): The foreground (text) color.
            bg (Optional[str]): The background color. Uses system default if None.
            bold (bool): Whether the button text should be bold.
            height (int): The height of the button in lines of text.
        
        Returns:
            tk.Button: The configured button widget.
        """
        weight = "bold" if bold else "normal"
        kwargs = {"text": text, "command": command, "fg": fg, "font": ("Arial", 10, weight), "height": height}
        if bg:
            kwargs["bg"] = bg
        return tk.Button(parent, **kwargs)

    # ==================== Main Window ====================

    def build_main_window(self) -> None:
        """
        Construct the primary application interface.
        
        Instantiates inputs for Excel files, text files, the text pasting area,
        and the main interaction buttons required to trigger a search.
        """
        main_pad = tk.Frame(self.root, padx=25, pady=20)
        main_pad.pack(fill="both", expand=True)
        
        # Spreadsheet Input Section
        self._create_label(main_pad, "Spreadsheet", font_size=11, bold=True).pack(pady=(0, 5))
        excel_frame = tk.Frame(main_pad)
        excel_frame.pack(fill="x", pady=(0, 15))
        self._create_button(excel_frame, "Select file.xlsx", 
                           command=lambda: self.select_file(self.excel_input_var, ("Excel files", "*.xlsx"))).pack(side="left", padx=(0, 10))
        tk.Entry(excel_frame, textvariable=self.excel_input_var, font=("Arial", 10), width=40).pack(side="left", fill="x", expand=True, ipady=5)

        # List/Search Items Input Section
        self._create_label(main_pad, "List", font_size=11, bold=True).pack(pady=(0, 5))
        list_frame = tk.Frame(main_pad)
        list_frame.pack(fill="x", pady=(0, 15))
        self._create_button(list_frame, "Select file.txt", 
                           command=lambda: self.select_file(self.txt_path_var, ("Text files", "*.txt"), True)).pack(side="left", padx=(0, 10))
        tk.Entry(list_frame, textvariable=self.txt_path_var, font=("Arial", 10), width=40).pack(side="left", fill="x", expand=True, ipady=5)

        # Direct text pasting area
        self.input_area = scrolledtext.ScrolledText(main_pad, height=5, width=65)
        self.input_area.pack(fill="y", expand=True, pady=10)

        # Primary Interaction
        self.check_button = self._create_button(main_pad, "Check Similarities", 
                                               command=self.run_similarity_check, 
                                               fg="#2980b9", height=2)
        self.check_button.pack(pady=10)

        self.progress_bar = self._create_progress_bar(main_pad)
        self._show_progress_bar(self.progress_bar)
        self._hide_progress_bar(self.progress_bar)

        self._create_label(main_pad, "Cabelo LTDA", font_size=8, bold=True, fg="gray").pack(side="bottom")

    # ==================== File & Input Handling ====================

    def select_file(self, target_var: tk.StringVar, file_types: Tuple[str, str], is_txt: bool = False) -> None:
        """
        Open a native OS file dialog to handle file selection.
        
        Args:
            target_var (tk.StringVar): The tkinter variable mapped to the entry field.
            file_types (Tuple[str, str]): Filter constraints for the file dialog.
            is_txt (bool): If True, reads the selected file's contents into the text area.
        """
        path = filedialog.askopenfilename(filetypes=[file_types])
        if path:
            target_var.set(path)
            if is_txt and self.input_area:
                with open(path, "r", encoding="utf-8") as f:
                    self.input_area.delete("1.0", tk.END)
                    self.input_area.insert(tk.END, f.read())

    # ==================== Similarity Checking ====================

    def run_similarity_check(self) -> None:
        """
        Validate inputs and initiate the fuzzy matching process safely.
        
        Spawns a background thread to prevent UI freezing during computation.
        Enforces a processing flag to disallow concurrent run clicks.
        """
        if self.processing: 
            return
        
        e_path = self.excel_input_var.get()
        t_content = self.input_area.get("1.0", tk.END).strip() if self.input_area else ""
        
        if not e_path or not t_content:
            messagebox.showwarning("Warning", "Input data missing.")
            return
        
        self.processing = True
        self.check_button.config(text="Checking...", state="disabled")
        self._show_progress_bar(self.progress_bar)
        threading.Thread(target=self._exec_check, args=(e_path, t_content), daemon=True).start()

    def _exec_check(self, path: str, content: str) -> None:
        """
        Execute the fuzzy matching operation inside the background thread.
        
        Args:
            path (str): The file path to the user's Excel data source.
            content (str): The raw text string containing items to search for.
        """
        try:
            df = load_excel_data(path)
            items = parse_search_items(content)
            pool = build_excel_pool(df)
            
            self.progress_bar['maximum'] = len(items)
            
            # Utilizing root.after for thread-safe UI updates
            results = find_smart_matches(
                items, pool, 
                progress_callback=lambda c, t: self.root.after(0, lambda: self.progress_bar.config(value=c))
            )
            self.root.after(0, lambda: self._on_check_done(results, df))
            
        except Exception as e:
            self.root.after(0, lambda err=e: [messagebox.showerror("Error", str(err)), self._reset_ui()])

    def _on_check_done(self, results: List[MultiMatchResult], df: pd.DataFrame) -> None:
        """
        Handle the successful completion of the background matching thread.
        
        Args:
            results (List[MultiMatchResult]): The processed match data.
            df (pd.DataFrame): The original loaded dataframe, passed forward for date extraction.
        """
        self._reset_ui()
        self.open_review_window(results, df)

    def _reset_ui(self) -> None:
        """Restore the primary interface elements to their default active state."""
        self.processing = False
        self.check_button.config(state="normal", text="Check Similarities")
        self._hide_progress_bar(self.progress_bar)

    # ==================== Review Window ====================

    def open_review_window(self, results: List[MultiMatchResult], df: pd.DataFrame) -> None:
        """
        Construct and present the user review window.
        
        Maintains the exact list order provided by the user. Assigns color coding
        to item headers based purely on the highest available match score for that item.
        The UI elements are centralized and the real-time blue selection data is removed.
        
        Args:
            results (List[MultiMatchResult]): The fuzzy match output data.
            df (pd.DataFrame): The original DataFrame used to retrieve final dates.
        """
        rw = Toplevel(self.root)
        rw.title("Review Matches")
        rw.withdraw()
        self.set_window_geometry(rw, width_pct=0.45, height_pct=0.8, min_w=600, min_h=500)

        # Retain original list order, no sorting applied
        review_items = results 

        def on_close() -> None:
            rw.unbind_all("<MouseWheel>")
            rw.destroy()
            
        rw.protocol("WM_DELETE_WINDOW", on_close)
        
        container = tk.Frame(rw)
        container.pack(fill="both", expand=True, padx=10, pady=10)
        
        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas)
        
        canvas_window = canvas.create_window((0,0), window=scroll_frame, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_window, width=e.width))
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Mousewheel binding
        rw.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        selection_map = {}
        
        for res in review_items:
            top_score = res.candidates[0].score if res.candidates else 0.0

            item_container = tk.Frame(scroll_frame, bd=1, relief="groove", pady=8)
            item_container.pack(fill="x", pady=4, padx=5)
            
            # Logic: Auto-select top candidate if score >= 69%
            is_low_conf = top_score < 0.69
            initial_val = "NONE" if (is_low_conf or not res.candidates) else res.candidates[0].suggested
            choice_var = tk.StringVar(value=initial_val)
            
            # Determine text color based on user score thresholds
            if top_score < 0.50:
                color = "#c0392b"  # Red (0 - 49%)
            elif top_score < 0.75:
                color = "#d35400"  # Orange (50 - 74%)
            elif top_score < 0.90:
                color = "#d4ac0d"  # Golden/Dark Yellow (75 - 89%) for better GUI contrast
            else:
                color = "#27ae60"  # Green (90 - 100%)

            # Centralized original item header
            self._create_label(item_container, f"• {res.original}", font_size=11, bold=True, fg=color).pack(anchor="center", pady=(0, 5))

            # Centralized options frame
            opts_frame = tk.Frame(item_container, bg="#fcfcfc")
            opts_frame.pack(pady=5)

            # Centralized Option to skip item entirely
            tk.Radiobutton(opts_frame, text="Skip", variable=choice_var, value="NONE", bg="#fcfcfc").pack(anchor="center", pady=2)

            # Centralized Candidates
            for c in res.candidates:
                tk.Radiobutton(opts_frame, text=f"[{int(c.score*100)}%] {c.suggested}", 
                              variable=choice_var, value=c.suggested, bg="#fcfcfc").pack(anchor="center", pady=2)
            
            selection_map[res.original] = choice_var

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        footer = tk.Frame(rw, height=80)
        footer.pack(side="bottom", fill="x")
        progress_bar = self._create_progress_bar(footer)
        
        def run_confirm() -> None:
            """Lock selections and process chosen data through a background thread."""
            confirm_btn.config(state="disabled")
            self._show_progress_bar(progress_bar)
            threading.Thread(target=self._final_task, args=(selection_map, results, df, rw, progress_bar), daemon=True).start()

        confirm_btn = self._create_button(footer, "Confirm Selections", command=run_confirm, fg='#27ae60', height=2)
        confirm_btn.pack(pady=10)
        rw.deiconify()

    # ==================== Final Processing ====================

    def _final_task(self, selection_map: dict, original_results: List[MultiMatchResult], df: pd.DataFrame, win: Toplevel, progress_bar: ttk.Progressbar) -> None:
        """
        Process the user's final selections and extract relative dates.
        
        Args:
            selection_map (dict): Dictionary linking original terms to tkinter StringVar choices.
            original_results (List[MultiMatchResult]): Baseline result map.
            df (pd.DataFrame): DataFrame for querying exact dates.
            win (Toplevel): Review window instance to destroy upon completion.
            progress_bar (ttk.Progressbar): The active progress bar instance.
        """
        progress_bar['maximum'] = len(original_results)
        final_results = []
        
        # Sequentially map over original_results to preserve ordering
        for i, res in enumerate(original_results):
            choice_var = selection_map.get(res.original)
            choice = choice_var.get() if choice_var else "NONE"
            
            match_display = choice if choice != "NONE" else "---"
            date = extract_dates_for_match(df, choice) if choice != "NONE" else "FOTO"
            
            final_results.append((res.original, match_display, date))
            self.root.after(0, lambda v=i+1: progress_bar.config(value=v))
        
        self.root.after(0, lambda: win.unbind_all("<MouseWheel>"))
        self.root.after(0, lambda: [self.show_results(final_results), win.destroy()]) 
    
    # ==================== Results Display ====================

    def show_results(self, data: List[Tuple[str, str, str]]) -> None:
        """
        Display final results in a formatted data table.
        
        Includes the functionality to bulk-copy all matched date results to clipboard.
        
        Args:
            data (List[Tuple[str, str, str]]): List containing tuples of 
                (original item, matched item, final date).
        """
        res_win = Toplevel(self.root)
        res_win.title("Fuzzy Rabbit - Results")
        self.set_window_geometry(res_win, width_pct=0.75, height_pct=0.7)

        self._create_label(res_win, "RESULTS", font_size=16, bold=True).pack(pady=10)

        table_frame = tk.Frame(res_win)
        table_frame.pack(fill="both", expand=True, padx=20, pady=10)

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
        
        style = ttk.Style()
        style.configure("Treeview", font=("Arial", 10), rowheight=28)
        tree.tag_configure('oddrow', background='#f7f7f7')
        tree.tag_configure('evenrow', background='white')

        for i, (orig, match, date) in enumerate(data):
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            tree.insert("", "end", values=(orig, match, date), tags=(tag,))

        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        footer = tk.Frame(res_win, pady=20)
        footer.pack(fill="x")

        def copy_dates_bt() -> None:
            """Isolate the date column, append to the clipboard, and update button state."""
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
    FuzzyFinderApp().root.mainloop()