import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import phonenumbers
from phonenumbers import NumberParseException
import csv
import os
import re

# ─── Optional: Excel support ──────────────────────────────────────
try:
    from openpyxl import load_workbook
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    load_workbook = None

try:
    import xlrd
    XLDR_AVAILABLE = True
except ImportError:
    XLDR_AVAILABLE = False

# ─── Validator ────────────────────────────────────────────────────

def validate_nigerian_number(number):
    try:
        number = number.strip()
        parsed = phonenumbers.parse(number, "NG")
        if (
            phonenumbers.is_valid_number(parsed)
            and phonenumbers.region_code_for_number(parsed) == "NG"
        ):
            normalised = phonenumbers.format_number(
                parsed,
                phonenumbers.PhoneNumberFormat.E164
            ).replace("+", "")
            return True, normalised
        return False, number
    except NumberParseException:
        return False, number


def read_csv(file_path):
    """Read CSV and return list of rows (each row as dict)."""
    rows = []
    with open(file_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return rows


def read_excel(file_path):
    """Read Excel (.xlsx or .xls) using available library."""
    ext = os.path.splitext(file_path)[1].lower()
    rows = []

    if ext == '.xlsx':
        if not OPENPYXL_AVAILABLE:
            raise ImportError("openpyxl not installed. Please run: pip install openpyxl")
        wb = load_workbook(file_path, data_only=True)
        ws = wb.active
        headers = [cell.value for cell in ws[1]]  # first row as headers
        for row in ws.iter_rows(min_row=2, values_only=True):
            if any(cell is not None for cell in row):  # skip empty rows
                row_dict = dict(zip(headers, row))
                rows.append(row_dict)

    elif ext == '.xls':
        if not XLDR_AVAILABLE:
            raise ImportError("xlrd not installed. Please run: pip install xlrd")
        wb = xlrd.open_workbook(file_path)
        sheet = wb.sheet_by_index(0)
        headers = [sheet.cell_value(0, col) for col in range(sheet.ncols)]
        for row_idx in range(1, sheet.nrows):
            row_values = [sheet.cell_value(row_idx, col) for col in range(sheet.ncols)]
            if any(v for v in row_values):
                row_dict = dict(zip(headers, row_values))
                rows.append(row_dict)

    return rows


def detect_phone_column(columns):
    """Return the column name that looks like 'phone' or None."""
    for col in columns:
        if re.search(r'phone|number|mobile|contact|telephone', str(col), re.IGNORECASE):
            return col
    return None


def load_numbers_from_file(file_path):
    """Extract numbers from file and insert into input_text."""
    ext = os.path.splitext(file_path)[1].lower()
    rows = []

    try:
        if ext == '.csv':
            rows = read_csv(file_path)
        elif ext in ('.xlsx', '.xls'):
            rows = read_excel(file_path)
        else:
            raise ValueError("Unsupported file type")

        if not rows:
            messagebox.showwarning("Empty Data", "File has no data rows.")
            return

        # Get column names from first row
        columns = list(rows[0].keys())
        phone_col = detect_phone_column(columns)

        if phone_col is None:
            # Prompt user
            selected = simpledialog.askstring(
                "Select Column",
                "Could not auto-detect a phone column.\n"
                "Enter the exact column name containing phone numbers:\n\n" +
                "\n".join(columns)
            )
            if not selected or selected not in columns:
                return
            phone_col = selected

        # Extract numbers
        numbers = []
        for row in rows:
            val = row.get(phone_col)
            if val is not None and str(val).strip():
                numbers.append(str(val).strip())

        if not numbers:
            messagebox.showwarning("No Numbers", f"Column '{phone_col}' has no phone numbers.")
            return

        # Populate input area
        input_text.delete("1.0", tk.END)
        input_text.insert(tk.END, "\n".join(numbers))
        status_var.set(f"Loaded {len(numbers)} numbers from {phone_col}")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to read file:\n{str(e)}")
        status_var.set("Error loading file")


# ─── UI Callbacks ─────────────────────────────────────────────────

def validate_numbers():
    valid_output.delete("1.0", tk.END)
    invalid_output.delete("1.0", tk.END)
    status_var.set("Validating...")
    root.update_idletasks()

    raw = input_text.get("1.0", tk.END).strip()
    if not raw:
        messagebox.showwarning("No Numbers", "Please paste numbers or upload a file first.")
        status_var.set("Ready")
        return

    numbers = [line.strip() for line in raw.splitlines() if line.strip()]
    valid_list, invalid_list = [], []

    for num in numbers:
        is_valid, result = validate_nigerian_number(num)
        if is_valid:
            valid_list.append(result)
        else:
            invalid_list.append(num)

    valid_output.insert(tk.END, "\n".join(valid_list) if valid_list else "None")
    invalid_output.insert(tk.END, "\n".join(invalid_list) if invalid_list else "None")

    total = len(numbers)
    valid_count = len(valid_list)
    invalid_count = len(invalid_list)
    stat_label.config(
        text=f"Total: {total}  |  Valid: {valid_count}  |  Invalid: {invalid_count}"
    )
    status_var.set(f"Done – {valid_count} valid, {invalid_count} invalid")


def clear_all():
    input_text.delete("1.0", tk.END)
    valid_output.delete("1.0", tk.END)
    invalid_output.delete("1.0", tk.END)
    stat_label.config(text="Total: 0  |  Valid: 0  |  Invalid: 0")
    status_var.set("Ready")
    file_label.config(text="No file selected")


def copy_output(text_widget, label):
    content = text_widget.get("1.0", tk.END).strip()
    if content and content != "None":
        root.clipboard_clear()
        root.clipboard_append(content)
        status_var.set(f"Copied {label} to clipboard")
    else:
        status_var.set("Nothing to copy")


def upload_file():
    file_path = filedialog.askopenfilename(
        title="Select file",
        filetypes=[
            ("Excel files", "*.xlsx *.xls"),
            ("CSV files", "*.csv"),
            ("All files", "*.*")
        ]
    )
    if not file_path:
        return

    file_label.config(text=os.path.basename(file_path))
    status_var.set(f"Reading {os.path.basename(file_path)}...")
    root.update_idletasks()
    load_numbers_from_file(file_path)


# ─── GUI ──────────────────────────────────────────────────────────

root = tk.Tk()
root.title("Nigerian Phone Number Validator")
root.geometry("900x750")
root.resizable(False, False)

style = ttk.Style()
style.theme_use("clam")

bg_colour = "#f5f7fa"
input_bg = "#ffffff"
accent = "#2b7a62"
font_family = "Segoe UI, Arial, sans-serif"

root.configure(bg=bg_colour)

# Header
header = tk.Frame(root, bg=accent, height=60)
header.pack(fill="x", pady=(0, 10))
header.pack_propagate(False)
tk.Label(
    header,
    text="📞 Nigerian Phone Number Validator",
    font=(font_family, 18, "bold"),
    bg=accent,
    fg="white"
).pack(expand=True)

# Main frame
main_frame = tk.Frame(root, bg=bg_colour)
main_frame.pack(fill="both", expand=True, padx=20, pady=10)

# Input
tk.Label(
    main_frame,
    text="Paste numbers (one per line) or upload a file",
    font=(font_family, 10, "bold"),
    bg=bg_colour,
    fg="#333"
).grid(row=0, column=0, sticky="w", pady=(0, 4))

input_text = tk.Text(
    main_frame,
    height=10,
    font=(font_family, 10),
    bg=input_bg,
    relief="flat",
    borderwidth=2,
    highlightthickness=2,
    highlightbackground="#d0d7de",
    highlightcolor=accent,
    padx=8,
    pady=8
)
input_text.grid(row=1, column=0, sticky="ew", pady=(0, 10))

# Buttons
btn_frame = tk.Frame(main_frame, bg=bg_colour)
btn_frame.grid(row=2, column=0, sticky="ew", pady=5)

tk.Button(
    btn_frame,
    text="✅ Validate",
    command=validate_numbers,
    bg=accent,
    fg="white",
    font=(font_family, 10, "bold"),
    padx=20,
    pady=6,
    relief="flat",
    cursor="hand2",
    activebackground="#1f5e4a",
    activeforeground="white"
).pack(side="left", padx=(0, 10))

tk.Button(
    btn_frame,
    text="🗑️ Clear All",
    command=clear_all,
    bg="#e74c3c",
    fg="white",
    font=(font_family, 10, "bold"),
    padx=20,
    pady=6,
    relief="flat",
    cursor="hand2",
    activebackground="#c0392b",
    activeforeground="white"
).pack(side="left", padx=(0, 10))

tk.Button(
    btn_frame,
    text="📂 Upload File",
    command=upload_file,
    bg="#2980b9",
    fg="white",
    font=(font_family, 10, "bold"),
    padx=20,
    pady=6,
    relief="flat",
    cursor="hand2",
    activebackground="#1a5276",
    activeforeground="white"
).pack(side="left")

file_label = tk.Label(
    btn_frame,
    text="No file selected",
    font=(font_family, 9),
    bg=bg_colour,
    fg="#555",
    anchor="w"
)
file_label.pack(side="left", padx=(15, 0))

# Stats
stat_label = tk.Label(
    main_frame,
    text="Total: 0  |  Valid: 0  |  Invalid: 0",
    font=(font_family, 10, "bold"),
    bg=bg_colour,
    fg="#333"
)
stat_label.grid(row=3, column=0, sticky="w", pady=(10, 6))

# Results
results_frame = tk.Frame(main_frame, bg=bg_colour)
results_frame.grid(row=4, column=0, sticky="ew", pady=5)

# Valid column
valid_col = tk.Frame(results_frame, bg=bg_colour)
valid_col.pack(side="left", fill="both", expand=True, padx=(0, 10))

tk.Label(
    valid_col,
    text="✅ Valid Numbers (234XXXXXXXXXX)",
    font=(font_family, 9, "bold"),
    bg=bg_colour,
    fg="#27ae60"
).pack(anchor="w")

valid_output = tk.Text(
    valid_col,
    height=10,
    font=(font_family, 9),
    bg=input_bg,
    relief="flat",
    borderwidth=2,
    highlightthickness=2,
    highlightbackground="#d0d7de",
    highlightcolor=accent,
    padx=6,
    pady=6,
    wrap="none"
)
valid_output.pack(fill="both", expand=True, pady=(2, 4))

tk.Button(
    valid_col,
    text="📋 Copy",
    command=lambda: copy_output(valid_output, "valid numbers"),
    bg="#27ae60",
    fg="white",
    font=(font_family, 9),
    padx=10,
    relief="flat",
    cursor="hand2",
    activebackground="#1e8449",
    activeforeground="white"
).pack(anchor="e")

# Invalid column
invalid_col = tk.Frame(results_frame, bg=bg_colour)
invalid_col.pack(side="left", fill="both", expand=True)

tk.Label(
    invalid_col,
    text="❌ Invalid Numbers",
    font=(font_family, 9, "bold"),
    bg=bg_colour,
    fg="#e74c3c"
).pack(anchor="w")

invalid_output = tk.Text(
    invalid_col,
    height=10,
    font=(font_family, 9),
    bg=input_bg,
    relief="flat",
    borderwidth=2,
    highlightthickness=2,
    highlightbackground="#d0d7de",
    highlightcolor=accent,
    padx=6,
    pady=6,
    wrap="none"
)
invalid_output.pack(fill="both", expand=True, pady=(2, 4))

tk.Button(
    invalid_col,
    text="📋 Copy",
    command=lambda: copy_output(invalid_output, "invalid numbers"),
    bg="#e74c3c",
    fg="white",
    font=(font_family, 9),
    padx=10,
    relief="flat",
    cursor="hand2",
    activebackground="#c0392b",
    activeforeground="white"
).pack(anchor="e")

# Status bar
status_var = tk.StringVar()
status_var.set("Ready")
tk.Label(
    root,
    textvariable=status_var,
    font=(font_family, 9),
    bg="#e9ecef",
    fg="#555",
    anchor="w",
    padx=10,
    pady=4,
    relief="sunken"
).pack(fill="x", side="bottom")

main_frame.columnconfigure(0, weight=1)
results_frame.columnconfigure(0, weight=1)
results_frame.columnconfigure(1, weight=1)

root.mainloop()