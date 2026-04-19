"""
Fuzzy Rabbit - Collapsible Review UI
- Items show only the best match by default.
- Click "▼ More" to expand alternatives for low-confidence matches.
- Radio buttons are tucked away inside the expandable section.
"""

from dataclasses import dataclass
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, Toplevel
from tkinter import ttk
from difflib import SequenceMatcher
import threading

@dataclass
class MatchCandidate:
    suggested: str
    score: float

@dataclass
class MultiMatchResult:
    original: str
    candidates: list[MatchCandidate]

def get_similarity(a, b):
    return SequenceMatcher(None, str(a).lower(), str(b).lower()).ratio()

def load_excel_data(path):
    return pd.read_excel(path)

def parse_search_items(text):
    return [line.strip() for line in text.splitlines() if line.strip()]

def build_excel_pool(df):
    return [str(x).strip() for x in pd.unique(df.values.ravel()) if pd.notna(x)]

def find_smart_matches(search_items, excel_pool, threshold=0.4, progress_callback=None):
    results = []
    for i, item in enumerate(search_items):
        all_scores = []
        for excel_value in excel_pool:
            score = get_similarity(item, excel_value)
            if score >= threshold:
                all_scores.append(MatchCandidate(suggested=excel_value, score=score))
        
        all_scores.sort(key=lambda x: x.score, reverse=True)
        
        # If best match >= 90%, we only really need that one, 
        # but we'll allow expansion if there are others.
        results.append(MultiMatchResult(original=item, candidates=all_scores[:3]))
            
        if progress_callback:
            progress_callback(i + 1, len(search_items))
    return results

def extract_dates_for_match(df, match_value):
    if not match_value or match_value == "NONE": return "No date found"
    match_str = str(match_value).strip()
    found_date_objs = []
    for col_name in df.columns:
        if df[col_name].astype(str).str.strip().eq(match_str).any():
            try:
                dt = pd.to_datetime(col_name)
                found_date_objs.append(dt)
            except: continue
    found_date_objs.sort(reverse=True)
    top_3 = [str(d.date()) for d in found_date_objs[:3]]
    return ", ".join(top_3) if top_3 else "No date found"

class FuzzyFinderApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Fuzzy Rabbit")
        self.root.minsize(550, 650)
        self.excel_input_var = tk.StringVar()
        self.txt_path_var = tk.StringVar()
        self.input_area = None
        self.processing = False
        self.build_main_window()

    def build_main_window(self):
        main_pad = tk.Frame(self.root, padx=25, pady=20)
        main_pad.pack(fill="both", expand=True)

        tk.Label(main_pad, text="Spreadsheet", font=("Arial", 11, "bold")).pack(pady=(0, 5))
        tk.Button(main_pad, text="Select File.xlsx", command=lambda: self.select_file(self.excel_input_var, ("Excel files", "*.xlsx"))).pack()
        tk.Entry(main_pad, textvariable=self.excel_input_var, width=60).pack(pady=(5, 15))

        tk.Label(main_pad, text="Comparison List", font=("Arial", 11, "bold")).pack(pady=(0, 5))
        tk.Button(main_pad, text="Select File.txt", command=lambda: self.select_file(self.txt_path_var, ("Text files", "*.txt"), True)).pack()
        tk.Entry(main_pad, textvariable=self.txt_path_var, width=60).pack(pady=(5, 15))

        self.input_area = scrolledtext.ScrolledText(main_pad, height=12, width=60)
        self.input_area.pack(fill="both", expand=True, pady=10)

        self.check_button = tk.Button(main_pad, text="Check Similarities", command=self.run_similarity_check, 
                                     fg="#2980b9", font=("Arial", 10, "bold"), height=2, width=25)
        self.check_button.pack(pady=10)

        self.progress_bar = ttk.Progressbar(main_pad, orient="horizontal", length=400, mode="determinate")
        self.progress_bar.pack(pady=5)
        self.progress_bar.pack_forget()

        tk.Label(main_pad, text="Cabelo LTDA", font=("Arial", 8, "bold"), fg="gray").pack(side="bottom")

    def select_file(self, target_var, file_types, is_txt=False):
        path = filedialog.askopenfilename(filetypes=[file_types])
        if path:
            target_var.set(path)
            if is_txt:
                with open(path, "r", encoding="utf-8") as f:
                    self.input_area.delete("1.0", tk.END)
                    self.input_area.insert(tk.END, f.read())

    def run_similarity_check(self):
        if self.processing: return
        e_path, t_content = self.excel_input_var.get(), self.input_area.get("1.0", tk.END).strip()
        if not e_path or not t_content:
            messagebox.showwarning("Warning", "Input data missing.")
            return
        self.processing = True
        self.check_button.config(state="disabled")
        self.progress_bar.pack(pady=10)
        threading.Thread(target=self._exec_check, args=(e_path, t_content), daemon=True).start()

    def _exec_check(self, path, content):
        try:
            df = load_excel_data(path)
            items = parse_search_items(content)
            pool = build_excel_pool(df)
            self.progress_bar['maximum'] = len(items)
            results = find_smart_matches(items, pool, progress_callback=lambda c, t: self.root.after(0, lambda: self.progress_bar.config(value=c)))
            self.root.after(0, lambda: self._on_check_done(results, df))
        except Exception as e:
            self.root.after(0, lambda: [messagebox.showerror("Error", str(e)), self._reset_ui()])

    def _on_check_done(self, results, df):
        self._reset_ui()
        results.sort(key=lambda x: x.candidates[0].score if x.candidates else 0)
        self.open_review_window(results, df)

    def _reset_ui(self):
        self.processing = False
        self.check_button.config(state="normal")
        self.progress_bar.pack_forget()

    def open_review_window(self, results, df):
        rw = Toplevel(self.root)
        rw.title("Review Matches")
        rw.geometry("900x750")
        
        def on_close():
            rw.unbind_all("<MouseWheel>")
            rw.destroy()
        rw.protocol("WM_DELETE_WINDOW", on_close)

        def _on_mousewheel(event):
            try: canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except: pass

        container = tk.Frame(rw)
        container.pack(fill="both", expand=True, padx=10, pady=10)
        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        rw.bind_all("<MouseWheel>", _on_mousewheel)

        selection_map = [] 

        for res in results:
            item_container = tk.Frame(scroll_frame, bd=1, relief="groove", pady=5)
            item_container.pack(fill="x", pady=2, padx=5)
            
            # Header Row
            header_row = tk.Frame(item_container)
            header_row.pack(fill="x", padx=5)

            choice_var = tk.StringVar()
            if res.candidates:
                choice_var.set(res.candidates[0].suggested)
            else:
                choice_var.set("NONE")

            # Main Label (The Search Term)
            tk.Label(header_row, text=f"• {res.original}", font=("Arial", 10, "bold"), width=30, anchor="w").pack(side="left")
            
            # Current Selection Label (Dynamic)
            current_sel_lbl = tk.Label(header_row, text=f"→ {choice_var.get()}", fg="#2980b9", font=("Arial", 9, "italic"))
            current_sel_lbl.pack(side="left", padx=10)

            # Expandable Section (Hidden by default)
            details_frame = tk.Frame(item_container, bg="#f9f9f9", pady=5)
            
            def toggle_details(f=details_frame, b=None):
                if f.winfo_viewable():
                    f.pack_forget()
                else:
                    f.pack(fill="x", padx=20)

            # Arrow Button for expansion
            expand_btn = tk.Button(header_row, text="▼ Alternatives", font=("Arial", 8), 
                                   command=toggle_details, relief="flat", fg="gray")
            expand_btn.pack(side="right")

            # Fill Expandable Section
            if res.candidates:
                tk.Radiobutton(details_frame, text="Skip / No Match", variable=choice_var, value="NONE", 
                               bg="#f9f9f9", command=lambda v=choice_var, l=current_sel_lbl: l.config(text=f"→ {v.get()}")).pack(anchor="w")
                
                for cand in res.candidates:
                    lbl_text = f"[{int(cand.score*100)}%] {cand.suggested}"
                    tk.Radiobutton(details_frame, text=lbl_text, variable=choice_var, value=cand.suggested, 
                                   bg="#f9f9f9", command=lambda v=choice_var, l=current_sel_lbl: l.config(text=f"→ {v.get()}")).pack(anchor="w")
            else:
                tk.Label(details_frame, text="No matches found.", fg="red", bg="#f9f9f9").pack(anchor="w")

            selection_map.append((res.original, choice_var))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        footer = tk.Frame(rw, height=110)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False) 
        prog = ttk.Progressbar(footer, orient="horizontal", length=400, mode="determinate")
        
        def run_final():
            btn.config(state="disabled")
            prog.pack(pady=10)
            threading.Thread(target=self._final_task, args=(selection_map, df, rw, prog), daemon=True).start()

        btn = tk.Button(footer, text="Confirm Selections", command=run_final, bg="#50906b", fg="white", font=("Arial", 10, "bold"), height=2, width=30)
        btn.pack(pady=10)

    def _final_task(self, selection_map, df, win, pbar):
        pbar['maximum'] = len(selection_map)
        final_results = []
        for i, (orig, var) in enumerate(selection_map):
            choice = var.get()
            if choice != "NONE":
                dates = extract_dates_for_match(df, choice)
                final_results.append((f"[{orig}] matched to '{choice}'", dates, True))
            else:
                final_results.append((f"{orig}: No match selected.", "No date found", False))
            self.root.after(0, lambda val=i+1: pbar.config(value=val))
        
        self.root.after(0, lambda: win.unbind_all("<MouseWheel>"))
        self.root.after(0, lambda: [self.show_results(final_results), win.destroy()])

    def show_results(self, data):
        data.sort(key=lambda x: (x[1] == "No date found"))
        res_win = Toplevel(self.root)
        res_win.title("Results")
        res_win.state('zoomed')
        txt = scrolledtext.ScrolledText(res_win, bg="#fdfdfd", font=("Courier New", 11), padx=15, pady=15)
        txt.pack(fill="both", expand=True)

        w_l, w_r = 130, 40
        txt.insert("end", f"{'MATCH IDENTIFIED':<{w_l}}{'3 MOST RECENT DATES':>{w_r}}\n" + ("=" * 170) + "\n\n")
        for item, dates, match in data:
            txt.insert("end", f"{item:<{w_l}}{dates:>{w_r}}\n\n")
        txt.config(state="disabled")

if __name__ == "__main__":
    FuzzyFinderApp().root.mainloop()