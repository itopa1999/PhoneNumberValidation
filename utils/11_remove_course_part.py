import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime
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
    purple = "#6f42c1"

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

    style.configure("Purple.TButton", background=purple, foreground="white")
    style.map("Purple.TButton",
              background=[("active", "#5a32a3"), ("pressed", "#4e2d8a")])

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
class AttendanceLogProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Attendance Log Processor")
        self.root.geometry("1000x700")
        self.root.configure(bg="#f0f4f8")

        self.style = setup_styles()

        self.file_path = tk.StringVar()
        self.data = None
        self.filtered_data = None
        self.original_count = 0
        self.filtered_count = 0

        main_frame = ttk.Frame(root, padding="15")
        main_frame.pack(fill="both", expand=True)

        # ----- Header / Title -----
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill="x", pady=(0, 10))

        title_label = ttk.Label(title_frame,
                                text="📊 Attendance Log Processor",
                                font=("Segoe UI", 18, "bold"),
                                foreground="#1a2b3c")
        title_label.pack(side="left")

        subtitle = ttk.Label(title_frame,
                             text="Filter and clean attendance logs",
                             font=("Segoe UI", 10),
                             foreground="#5a6b7c")
        subtitle.pack(side="left", padx=(15, 0))

        # ----- Description -----
        desc_text = (
            "📌 This app processes attendance log files.\n"
            "   • Select a CSV or Excel file with columns: Time, User full name, Event name.\n"
            "   • Click 'Process' to remove rows containing 'Course module viewed'.\n"
            "   • View the cleaned data in the table below.\n"
            "   • Save the filtered results as CSV or Excel using the download buttons."
        )
        desc_label = ttk.Label(main_frame, text=desc_text, font=("Segoe UI", 9),
                               foreground="#2a4b6c", background="#f0f4f8",
                               justify="left", wraplength=900)
        desc_label.pack(anchor="w", pady=(0, 10))

        ttk.Separator(main_frame, orient="horizontal").pack(fill="x", pady=(0, 12))

        # ----- File selection frame -----
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(file_frame, text="Log File:").pack(side="left", padx=(0, 8))

        entry = ttk.Entry(file_frame, textvariable=self.file_path, width=50, font=("Segoe UI", 10))
        entry.pack(side="left", padx=(0, 8), fill="x", expand=True)

        # Toolbar buttons
        browse_btn = ttk.Button(file_frame, text="📂 Browse", command=self.browse_file)
        browse_btn.pack(side="left", padx=(0, 6))

        process_btn = ttk.Button(file_frame, text="⚡ Process & Filter", command=self.process_file,
                                 style="Success.TButton")
        process_btn.pack(side="left", padx=(0, 6))

        clear_btn = ttk.Button(file_frame, text="🗑️ Clear", command=self.clear_all,
                               style="Danger.TButton")
        clear_btn.pack(side="left", padx=(0, 6))

        # ----- Status bar -----
        self.status = ttk.Label(main_frame, text="Ready", relief="sunken",
                                anchor="w", style="Status.TLabel")
        self.status.pack(fill="x", pady=(0, 10))

        # ----- Treeview -----
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tree_frame,
                                 columns=("Time", "User", "Event"),
                                 show="headings",
                                 height=20)

        self.tree.heading("Time", text="Time")
        self.tree.heading("User", text="User Full Name")
        self.tree.heading("Event", text="Event Name")

        self.tree.column("Time", width=200, anchor="center")
        self.tree.column("User", width=350, anchor="w")
        self.tree.column("Event", width=300, anchor="w")

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # ----- Bottom action bar -----
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill="x", pady=(12, 0))

        # Left side: Count info
        self.count_label = ttk.Label(bottom_frame, text="", font=("Segoe UI", 9, "italic"))
        self.count_label.pack(side="left")

        # Right side: Download buttons
        download_frame = ttk.Frame(bottom_frame)
        download_frame.pack(side="right")

        ttk.Button(download_frame, text="📥 Download CSV", 
                   command=lambda: self.save_result("csv"),
                   style="Primary.TButton").pack(side="left", padx=(0, 6))
        
        ttk.Button(download_frame, text="📥 Download Excel", 
                   command=lambda: self.save_result("excel"),
                   style="Purple.TButton").pack(side="left", padx=(0, 6))
        
        ttk.Button(download_frame, text="Quit", command=root.quit,
                   style="Danger.TButton").pack(side="left")

    # ---------- Methods ----------
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select Attendance Log File",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.file_path.set(filename)
            self.status.config(text=f"📁 Loaded: {os.path.basename(filename)}")

    def process_file(self):
        if not self.file_path.get():
            messagebox.showerror("Error", "Please select a log file first.")
            return

        try:
            # Load the file
            file_ext = os.path.splitext(self.file_path.get())[1].lower()
            if file_ext in ['.csv']:
                df = pd.read_csv(self.file_path.get())
            elif file_ext in ['.xlsx', '.xls']:
                df = pd.read_excel(self.file_path.get())
            else:
                messagebox.showerror("Error", "Unsupported file format. Please use CSV or Excel files.")
                return

            # Check required columns
            required_cols = ["Time", "User full name", "Event name"]
            if not all(col in df.columns for col in required_cols):
                messagebox.showerror("Error",
                                     f"File must contain columns: {', '.join(required_cols)}\n"
                                     f"Found: {', '.join(df.columns)}")
                return

            self.original_count = len(df)
            self.data = df.copy()

            # Filter out rows with "Course module viewed"
            filtered_df = df[~df["Event name"].str.contains("Course module viewed", case=False, na=False)]
            self.filtered_count = len(filtered_df)
            self.filtered_data = filtered_df.copy()

            # Populate treeview
            self.populate_tree(filtered_df)

            # Update status
            removed_count = self.original_count - self.filtered_count
            self.status.config(
                text=f"✅ Processed: {self.original_count} rows → {self.filtered_count} rows (removed {removed_count} 'Course module viewed' entries)"
            )
            self.count_label.config(
                text=f"📊 Original: {self.original_count} | Filtered: {self.filtered_count} | Removed: {removed_count}"
            )
            
            messagebox.showinfo("Success",
                                f"Filtering complete!\n"
                                f"Original rows: {self.original_count}\n"
                                f"Filtered rows: {self.filtered_count}\n"
                                f"Removed: {removed_count} 'Course module viewed' entries")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to process file:\n{str(e)}")
            self.status.config(text="❌ Error")

    def populate_tree(self, df):
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Insert filtered data
        for _, row in df.iterrows():
            self.tree.insert("", "end", values=[row["Time"], row["User full name"], row["Event name"]])

    def save_result(self, format_type):
        if self.filtered_data is None or self.filtered_data.empty:
            messagebox.showwarning("Warning", "No data to save. Process a file first.")
            return

        # Generate default filename
        base_name = "filtered_attendance_log"
        if format_type == "csv":
            file_types = [("CSV files", "*.csv")]
            default_ext = ".csv"
        else:
            file_types = [("Excel files", "*.xlsx")]
            default_ext = ".xlsx"

        file_path = filedialog.asksaveasfilename(
            title=f"Save Filtered Data as {format_type.upper()}",
            defaultextension=default_ext,
            filetypes=file_types,
            initialfile=f"{base_name}{default_ext}"
        )
        
        if file_path:
            try:
                if format_type == "csv":
                    self.filtered_data.to_csv(file_path, index=False)
                else:
                    self.filtered_data.to_excel(file_path, index=False)
                
                messagebox.showinfo("Success", f"File saved to:\n{file_path}")
                self.status.config(text=f"💾 Saved: {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not save file:\n{str(e)}")

    def clear_all(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.data = None
        self.filtered_data = None
        self.file_path.set("")
        self.original_count = 0
        self.filtered_count = 0
        self.status.config(text="Cleared")
        self.count_label.config(text="")

if __name__ == "__main__":
    root = tk.Tk()
    app = AttendanceLogProcessor(root)
    root.mainloop()