import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, Toplevel
from difflib import SequenceMatcher

def get_similarity(a, b):
    return SequenceMatcher(None, str(a).lower(), str(b).lower()).ratio()

def select_file(target_var, file_types, is_txt=False):
    path = filedialog.askopenfilename(filetypes=[file_types])
    if path:
        target_var.set(path)
        if is_txt:
            with open(path, 'r') as f:
                input_area.delete('1.0', tk.END)
                input_area.insert(tk.END, f.read())

def run_similarity_check():
    excel_input = excel_input_var.get()
    txt_input = input_area.get("1.0", tk.END).strip()

    if not excel_input or not txt_input:
        messagebox.showwarning("Missing Data", "Please select an Excel file and provide the searching items.")
        return

    try:
        # 1. Load Data
        df = pd.read_excel(excel_input)
        search_items = [line.strip() for line in txt_input.split('\n') if line.strip()]
        
        # 2. Get all unique values from Excel to compare against
        excel_pool = pd.unique(df.values.ravel())
        excel_pool = [str(x) for x in excel_pool if pd.notna(x)]

        # 3. Find Best Matches
        matches_to_review = []
        for item in search_items:
            best_val = None
            best_score = 0
            
            for excel_val in excel_pool:
                score = get_similarity(item, excel_val)
                if score > best_score:
                    best_score = score
                    best_val = excel_val
            
            if best_score >= 0.6:
                matches_to_review.append({'original': item, 'suggested': best_val, 'score': best_score})
            else:
                matches_to_review.append({'original': item, 'suggested': "NONE FOUND", 'score': 0})

        open_review_window(matches_to_review, df)

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def open_review_window(matches, df):
    """Pop-up window for user to confirm or deny suggested matches."""
    review_win = Toplevel(root)
    review_win.title("Review & Confirm Matches")
    review_win.geometry("1000x600")

    tk.Label(review_win, text="Check the matches you want to include in the search:", font=("Arial", 18, "bold")).pack(pady=25)

    # Scrollable area for the list
    container = tk.Frame(review_win)
    container.pack(fill="both", expand=True, padx=40)
    canvas = tk.Canvas(container)
    scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    # Checkbox for every match
    check_vars = []

    for m in matches:
        var = tk.BooleanVar(value=(m['score'] > 0.9)) # Auto-check if very high confidence
        frame = tk.Frame(scrollable_frame)
        frame.pack(fill="x", anchor="w", pady=2)
        
        display_text = f"Input: '{m['original']}'  ->  Excel: '{m['suggested']}' ({int(m['score']*100)}%)"
        chk = tk.Checkbutton(frame, text=display_text, variable=var)
        chk.pack(side="left")
        check_vars.append((m, var))

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def process_final():
        res_win = Toplevel(root)
        res_win.title("Results")
        res_win.geometry("1400x900")

        tk.Label(res_win, text="Search Results", font=("Arial", 14, "bold")).pack(pady=20)

        results_display = scrolledtext.ScrolledText(res_win, width=100, height=90, bg="#ecf0f1")
        results_display.pack(padx=15, pady=20)

        # Run the loop and insert into results_display
        for m, var in check_vars:
            if var.get() and m['suggested'] != "NONE FOUND":
                formatted_cols = []
                for col in df.columns:
                    if m['suggested'] in df[col].astype(str).values:
                        if hasattr(col, 'strftime'):
                            formatted_cols.append(col.strftime('%Y-%m-%d'))
                        else:
                            formatted_cols.append(str(col))
                
                # Insert text into the widget we just created
                results_display.insert(tk.END, f"{m['original']} > '{m['suggested']}' located at {', '.join(formatted_cols)}\n")
            else:
                results_display.insert(tk.END, f"{m['original']}: No confirmed match.\n")
        
        results_display.config(state=tk.DISABLED)
        review_win.destroy()

    tk.Button(review_win, text="Confirm Review", command=process_final, bg="#50906b", fg="white", font=("Arial", 12, "bold")).pack(pady=30)

    """Separate window to display final results."""

# --- Main Window Setup ---

root = tk.Tk()
root.title("Fuzzy Rabbit")
root.geometry("600x800")

excel_input_var = tk.StringVar()
txt_path_var = tk.StringVar()

# Section 1: Excel File
tk.Label(root, text="Spreadsheet", font=("Arial", 14, "bold")).pack(pady=15)
tk.Button(root, text="Select File.xlsx", command=lambda: select_file(excel_input_var, ("Excel files", "*.xlsx"))).pack(pady=5)
tk.Entry(root, textvariable=excel_input_var, width=60).pack(padx=10, pady=10)

# Section 2: Comparison File
tk.Label(root, text="Comparison List", font=("Arial", 14, "bold")).pack(pady=15)
tk.Button(root, text="Select File.txt", command=lambda: select_file(txt_path_var, ("Text files", "*.txt"), True)).pack(pady=5)
tk.Entry(root, textvariable=txt_path_var, width=60).pack(padx=10, pady=10)

input_area = scrolledtext.ScrolledText(root, height=20, width=75)
input_area.pack(padx=35, pady=20)

# Section 3: Action
tk.Button(root, text="Check Similarities", command=run_similarity_check, fg="#2980b9", font=("Arial", 10, "bold"), height=2).pack(pady=20)

tk.Label(root, text="Cabelo LTDA", font=("Arial", 8, "bold")).pack(side="bottom", pady=5)

root.mainloop()