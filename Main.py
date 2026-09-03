import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import re
from tksheet import Sheet

# Global placeholders
df = pd.DataFrame()
history = []


def save_to_history():
    """Saves a deep copy of the current DataFrame state to the history stack."""
    global df, history
    history.append(df.copy())
    if len(history) > 20:
        history.pop(0)


def open_excel_file():
    """Opens Windows Explorer to select an Excel file and loads it into the grid."""
    global df, history
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv"), ("All Files", "*.*")]
    )
    if file_path:
        try:
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path).fillna("").astype(str)
            else:
                df = pd.read_excel(file_path).fillna("").astype(str)
            history.clear()
            sheet.set_sheet_data(df.values.tolist(), redraw=False)
            sheet.headers(list(df.columns))
            sheet.redraw()
        except Exception as e:
            messagebox.showerror("Error Loading File", f"Could not read the Excel file:\n{e}")


def export_data():
    """Exports the current DataFrame to an Excel or CSV file."""
    global df
    if df.empty:
        messagebox.showwarning("No Data", "There is no data to export! Please open a file first.")
        return

    file_path = filedialog.asksaveasfilename(
        title="Save Cleaned Data",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")]
    )

    if file_path:
        try:
            if file_path.endswith('.csv'):
                df.to_csv(file_path, index=False)
            else:
                df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", "Data exported successfully!")
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not save the file:\n{e}")


def parse_by_template(raw_value, mask):
    """
    Unified Smart Parser Engine with Escape Tokens:
    - /N: Extracts the next Number.
    - /L: Extracts the next Letter as UPPERCASE.
    - /l: Extracts the next Letter as lowercase.
    - /+: Dumps the rest of the numbers or letters depending on the last command.
    - /*: Dumps the exact remainder of the original raw string.
    - /-: Cuts off parsing here, deleting/ignoring everything after.
    - Any other character is treated as literal text.
    """
    if not mask:
        return raw_value

    digits = list(re.finditer(r'\d', raw_value))
    letters = list(re.finditer(r'[a-zA-Z]', raw_value))

    d_idx = 0
    l_idx = 0
    result = []

    i = 0
    last_type = None

    while i < len(mask):
        # Look for the slash trigger and valid command (including '-')
        if mask[i] == '/' and i + 1 < len(mask) and mask[i + 1] in ('N', 'L', 'l', '+', '*', '-'):
            cmd = mask[i + 1]
            i += 2  # Jump past the slash and the command letter

            if cmd == 'N':
                if d_idx < len(digits):
                    result.append(digits[d_idx].group())
                    d_idx += 1
                last_type = 'N'
            elif cmd == 'L':
                if l_idx < len(letters):
                    result.append(letters[l_idx].group().upper())
                    l_idx += 1
                last_type = 'L'
            elif cmd == 'l':
                if l_idx < len(letters):
                    result.append(letters[l_idx].group().lower())
                    l_idx += 1
                last_type = 'l'
            elif cmd == '+':
                if last_type == 'N':
                    result.append("".join([m.group() for m in digits[d_idx:]]))
                    d_idx = len(digits)
                elif last_type == 'L':
                    result.append("".join([m.group().upper() for m in letters[l_idx:]]))
                    l_idx = len(letters)
                elif last_type == 'l':
                    result.append("".join([m.group().lower() for m in letters[l_idx:]]))
                    l_idx = len(letters)
            elif cmd == '*':
                start_pos = 0
                if last_type == 'N' and d_idx < len(digits):
                    start_pos = digits[d_idx].start()
                elif last_type in ('L', 'l') and l_idx < len(letters):
                    start_pos = letters[l_idx].start()
                elif d_idx == len(digits) and l_idx == len(letters):
                    break

                result.append(raw_value[start_pos:])
                break
            elif cmd == '-':
                # /- Cuts off output here, deleting/ignoring everything after
                break
        else:
            result.append(mask[i])
            i += 1

    return "".join(result)


def apply_custom_formatting():
    global df
    if df.empty:
        messagebox.showwarning("No Data", "Please open an Excel file first!")
        return

    selected_cells = set(sheet.get_selected_cells())
    for box in sheet.get_all_selection_boxes():
        r1, c1, r2, c2 = box[0], box[1], box[2], box[3]
        for r in range(r1, r2):
            for c in range(c1, c2):
                selected_cells.add((r, c))

    active = sheet.get_currently_selected()
    if active and not selected_cells:
        r = active[0] if isinstance(active, tuple) else getattr(active, "row", None)
        c = active[1] if isinstance(active, tuple) else getattr(active, "column", None)
        if r is not None and c is not None:
            selected_cells.add((r, c))

    if not selected_cells:
        messagebox.showinfo("Selection Required", "Please select cells to format.")
        return

    chosen_format = format_combo.get()
    custom_mask = mask_entry.get().strip()

    save_to_history()

    for row, col in selected_cells:
        val = str(df.iloc[row, col]).strip()

        if chosen_format == "Clear NaN & Empty Cells":
            cleaned_value = "" if (re.match(r'^nan$', val, re.IGNORECASE) or val == "") else val
        elif chosen_format == "Capitalize Each Word (Title Case)":
            cleaned_value = val.title()
        elif chosen_format == "UPPERCASE":
            cleaned_value = val.upper()
        elif chosen_format == "lowercase":
            cleaned_value = val.lower()
        elif chosen_format == "Letters Only":
            cleaned_value = re.sub(r'[^a-zA-Z\s]', '', val)
        elif chosen_format == "Numbers Only":
            cleaned_value = re.sub(r'[^0-9]', '', val)
        elif chosen_format == "Use Custom Mask Below 👇":
            if not custom_mask:
                messagebox.showwarning("Empty Mask", "Please enter a valid format mask first!")
                return
            cleaned_value = parse_by_template(val, custom_mask)
        else:
            cleaned_value = val

        df.iloc[row, col] = cleaned_value
        sheet.set_cell_data(row, col, cleaned_value, redraw=False)

    sheet.redraw()


def toggle_mask_field(event=None):
    if format_combo.get() == "Use Custom Mask Below 👇":
        mask_entry.config(state="normal", bg="white")
    else:
        mask_entry.config(state="disabled", bg="#f0f0f0")


def undo_last_action():
    global df, history
    if not history:
        messagebox.showinfo("Undo", "Nothing to undo!")
        return
    df = history.pop()
    sheet.set_sheet_data(df.values.tolist(), redraw=False)
    sheet.redraw()


# --- UI Setup ---
root = tk.Tk()
root.title("Advanced Custom Template Data Engine V5")
root.geometry("1100x650")

# Toolbars
file_toolbar = tk.LabelFrame(root, text=" File Actions ", padx=10, pady=5)
file_toolbar.pack(fill="x", padx=15, pady=5)

format_toolbar = tk.LabelFrame(root, text=" Text Formatting Engine ", padx=10, pady=10)
format_toolbar.pack(fill="x", padx=15, pady=5)

# --- File Actions ---
tk.Button(file_toolbar, text="📂 Open Excel", command=open_excel_file, bg="#2196F3", fg="white",
          font=("Arial", 9, "bold")).pack(side="left", padx=5)

tk.Button(file_toolbar, text="💾 Export Data", command=export_data, bg="#9C27B0", fg="white",
          font=("Arial", 9, "bold")).pack(side="left", padx=5)

tk.Button(file_toolbar, text="↩️ Undo Last Action", command=undo_last_action, bg="#FF9800", fg="white",
          font=("Arial", 9, "bold")).pack(side="left", padx=20)

# --- Formatting Layout Elements ---
tk.Label(format_toolbar, text="Style Preset:", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky="w", padx=5,
                                                                               pady=2)

format_options = [
    "Clear NaN & Empty Cells",
    "Capitalize Each Word (Title Case)",
    "UPPERCASE",
    "lowercase",
    "Letters Only",
    "Numbers Only",
    "Use Custom Mask Below 👇"
]
format_combo = ttk.Combobox(format_toolbar, values=format_options, width=32, state="readonly")
format_combo.set("Clear NaN & Empty Cells")
format_combo.grid(row=0, column=1, padx=5, pady=2)
format_combo.bind("<<ComboboxSelected>>", toggle_mask_field)

# Custom Template Entry
tk.Label(format_toolbar, text="Custom Mask Layout:", font=("Arial", 9, "bold")).grid(row=1, column=0, sticky="w",
                                                                                     padx=5, pady=5)
mask_entry = tk.Entry(format_toolbar, width=35, font=("Courier New", 10), state="disabled", bg="#f0f0f0")
mask_entry.grid(row=1, column=1, padx=5, pady=5)

# Quick instructions label updated to reflect /-
hint_lbl = tk.Label(
    format_toolbar,
    text="(Mask codes: /N=Num | /L=Upper | /l=Lower | /+=Dump Rest | /*=Keep Raw | /-=Cut off)\n(Any text without a slash is printed exactly as typed)",
    fg="#666666",
    font=("Arial", 8, "italic"),
    justify="left"
)
hint_lbl.grid(row=1, column=2, padx=10, pady=5, sticky="w")

# Main action button execution
apply_btn = tk.Button(format_toolbar, text="🚀 Apply to Selected Cells", command=apply_custom_formatting, bg="#4CAF50",
                      fg="white", font=("Arial", 10, "bold"), padx=15)
apply_btn.grid(row=0, column=2, rowspan=1, padx=10, pady=2, sticky="ns")

# --- Spreadsheet Display ---
spreadsheet_frame = tk.Frame(root)
spreadsheet_frame.pack(expand=True, fill="both", padx=15, pady=10)

sheet = Sheet(spreadsheet_frame, data=[[]], headers=[])
sheet.enable_bindings()
sheet.pack(expand=True, fill="both")

root.mainloop()