import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import timedelta
import os

# ---------- Custom ttk styles ----------
def setup_styles():
    style = ttk.Style()
    style.theme_use('clam')

    bg = "#f0f4f8"
    fg = "#1a2b3c"
    accent = "#2a7de1"
    success = "#28a745"
    danger = "#dc3545"
    warning = "#ffc107"

    style.configure("TFrame", background=bg)
    style.configure("TLabel", background=bg, foreground=fg, font=("Segoe UI", 10))
    style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6)
    style.map("TButton",
              background=[("active", accent), ("pressed", accent)],
              foreground=[("active", "white"), ("pressed", "white")])

    style.configure("Success.TButton", background=success, foreground="white")
    style.map("Success.TButton",
              background=[("active", "#218838"), ("pressed", "#1e7e34")])

    style.configure("Primary.TButton", background=accent, foreground="white")
    style.map("Primary.TButton",
              background=[("active", "#1a5fb4"), ("pressed", "#15539e")])

    style.configure("Danger.TButton", background=danger, foreground="white")
    style.map("Danger.TButton",
              background=[("active", "#c82333"), ("pressed", "#bd2130")])

    style.configure("Treeview",
                    background="white",
                    foreground=fg,
                    fieldbackground="white",
                    font=("Segoe UI", 9))
    style.map("Treeview", background=[("selected", accent)])

    style.configure("Treeview.Heading",
                    background="#d9e2ec",
                    foreground=fg,
                    font=("Segoe UI", 10, "bold"))

    style.configure("Status.TLabel",
                    background="#e9ecef",
                    foreground=fg,
                    font=("Segoe UI", 9),
                    padding=4)

    return style

# ---------- Main Application ----------
class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Attendance Deduplicator")
        self.root.geometry("950x650")
        self.root.configure(bg="#f0f4f8")

        self.style = setup_styles()

        self.file_path = tk.StringVar()
        self.data = None
        self.result = None

        main_frame = ttk.Frame(root, padding="15")
        main_frame.pack(fill="both", expand=True)

        # ----- Header / Title -----
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill="x", pady=(0, 10))

        title_label = ttk.Label(title_frame,
                                text="📊 Attendance Deduplicator",
                                font=("Segoe UI", 18, "bold"),
                                foreground="#1a2b3c")
        title_label.pack(side="left")

        subtitle = ttk.Label(title_frame,
                             text="Combine duplicate entries and sum durations",
                             font=("Segoe UI", 10),
                             foreground="#5a6b7c")
        subtitle.pack(side="left", padx=(15, 0))

        # ----- Description -----
        desc_text = (
            "📌 This app deduplicates attendance records by student name.\n"
            "   • Select an Excel file with columns: Name, Join, Left, Duration.\n"
            "   • Click 'Process' – duplicate names are merged, and durations are summed.\n"
            "   • The output shows one row per student with total duration.\n"
            "   • Save the cleaned list as a new Excel file using 'Save as Excel'."
        )
        desc_label = ttk.Label(main_frame, text=desc_text, font=("Segoe UI", 9),
                               foreground="#2a4b6c", background="#f0f4f8",
                               justify="left", wraplength=900)
        desc_label.pack(anchor="w", pady=(0, 10))

        ttk.Separator(main_frame, orient="horizontal").pack(fill="x", pady=(0, 12))

        # ----- File selection frame -----
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(file_frame, text="Excel File:").pack(side="left", padx=(0, 8))

        entry = ttk.Entry(file_frame, textvariable=self.file_path, width=55, font=("Segoe UI", 10))
        entry.pack(side="left", padx=(0, 8), fill="x", expand=True)

        # Toolbar buttons
        browse_btn = ttk.Button(file_frame, text="📂 Browse", command=self.browse_file)
        browse_btn.pack(side="left", padx=(0, 6))

        process_btn = ttk.Button(file_frame, text="⚡ Process", command=self.process_file,
                                 style="Success.TButton")
        process_btn.pack(side="left", padx=(0, 6))

        clear_btn = ttk.Button(file_frame, text="🗑️ Clear", command=self.clear_all,
                               style="Danger.TButton")
        clear_btn.pack(side="left", padx=(0, 6))

        save_btn = ttk.Button(file_frame, text="💾 Save as Excel", command=self.save_result,
                              style="Primary.TButton")
        save_btn.pack(side="left")

        # ----- Status bar -----
        self.status = ttk.Label(main_frame, text="Ready", relief="sunken",
                                anchor="w", style="Status.TLabel")
        self.status.pack(fill="x", pady=(0, 10))

        # ----- Treeview -----
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tree_frame,
                                 columns=("Name", "Join", "Left", "Duration"),
                                 show="headings",
                                 height=18)

        self.tree.heading("Name", text="Name")
        self.tree.heading("Join", text="Join Time")
        self.tree.heading("Left", text="Left Time")
        self.tree.heading("Duration", text="Total Duration (HH:MM:SS)")

        self.tree.column("Name", width=280, anchor="w")
        self.tree.column("Join", width=180, anchor="center")
        self.tree.column("Left", width=180, anchor="center")
        self.tree.column("Duration", width=150, anchor="center")

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # ----- Bottom action bar -----
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill="x", pady=(12, 0))

        ttk.Button(bottom_frame, text="Quit", command=root.quit,
                   style="Danger.TButton").pack(side="right")

        self.count_label = ttk.Label(bottom_frame, text="", font=("Segoe UI", 9, "italic"))
        self.count_label.pack(side="right", padx=(0, 15))

    # ---------- Methods ----------
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select Attendance Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.file_path.set(filename)
            self.status.config(text=f"📁 Loaded: {os.path.basename(filename)}")

    def process_file(self):
        if not self.file_path.get():
            messagebox.showerror("Error", "Please select an Excel file first.")
            return

        try:
            df = pd.read_excel(self.file_path.get())
            required_cols = ["Name", "Join", "Left", "Duration"]
            if not all(col in df.columns for col in required_cols):
                messagebox.showerror("Error",
                                     f"Excel must contain columns: {', '.join(required_cols)}")
                return

            self.data = df.copy()

            # ---------- Parse Join and Left with explicit format ----------
            date_format = "%m/%d/%Y, %I:%M:%S %p"
            df["Join_dt"] = pd.to_datetime(df["Join"], format=date_format, errors='coerce')
            df["Left_dt"] = pd.to_datetime(df["Left"], format=date_format, errors='coerce')

            # Parse existing Duration
            def parse_duration(dur_str):
                if isinstance(dur_str, timedelta):
                    return dur_str
                try:
                    parts = str(dur_str).strip().split(":")
                    if len(parts) == 3:
                        h, m, s = map(int, parts)
                        return timedelta(hours=h, minutes=m, seconds=s)
                    elif len(parts) == 2:
                        m, s = map(int, parts)
                        return timedelta(minutes=m, seconds=s)
                    else:
                        return timedelta(0)
                except:
                    return timedelta(0)

            df["Duration_TD"] = df["Duration"].apply(parse_duration)

            # Validate and correct durations
            df["True_Duration_TD"] = df["Left_dt"] - df["Join_dt"]

            correction_count = 0
            for idx, row in df.iterrows():
                true_td = row["True_Duration_TD"]
                given_td = row["Duration_TD"]
                if pd.notna(true_td):
                    if abs((true_td - given_td).total_seconds()) > 1:
                        df.at[idx, "Duration_TD"] = true_td
                        correction_count += 1

            if correction_count > 0:
                self.status.config(text=f"⚠️ Corrected {correction_count} duration mismatches")
            else:
                self.status.config(text="✅ All durations are consistent")

            # Group by Name
            grouped = df.groupby("Name", as_index=False).agg({
                "Join": "first",
                "Left": "last",
                "Duration_TD": "sum"
            })

            # Format duration as "H:MM:SS" (no leading zero for hours)
            def format_timedelta(td):
                total_seconds = int(td.total_seconds())
                hours, rem = divmod(total_seconds, 3600)
                minutes, seconds = divmod(rem, 60)
                return f"{hours}:{minutes:02d}:{seconds:02d}"

            grouped["Duration"] = grouped["Duration_TD"].apply(format_timedelta)
            grouped.drop("Duration_TD", axis=1, inplace=True)
            grouped = grouped[["Name", "Join", "Left", "Duration"]]

            self.result = grouped
            self.populate_tree(grouped)

            self.status.config(
                text=f"✅ Processed: {len(df)} rows → {len(grouped)} unique names"
            )
            self.count_label.config(text=f"{len(grouped)} unique entries")
            messagebox.showinfo("Success",
                                f"Deduplication complete.\nOriginal: {len(df)} rows\nUnique: {len(grouped)} rows")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to process file:\n{str(e)}")
            self.status.config(text="❌ Error")

    def populate_tree(self, df):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def save_result(self):
        if self.result is None or self.result.empty:
            messagebox.showwarning("Warning", "No data to save. Process a file first.")
            return

        file_path = filedialog.asksaveasfilename(
            title="Save Deduplicated Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            try:
                self.result.to_excel(file_path, index=False)
                messagebox.showinfo("Success", f"File saved to:\n{file_path}")
                self.status.config(text=f"💾 Saved: {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not save file:\n{str(e)}")

    def clear_all(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.data = None
        self.result = None
        self.file_path.set("")
        self.status.config(text="Cleared")
        self.count_label.config(text="")

if __name__ == "__main__":
    root = tk.Tk()
    app = AttendanceApp(root)
    root.mainloop()