import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime, timedelta
import os
import re

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
    orange = "#fd7e14"

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

    style.configure("Orange.TButton", background=orange, foreground="white")
    style.map("Orange.TButton",
              background=[("active", "#e06b0a"), ("pressed", "#cc5f09")])

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
class AttendanceDurationCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("Meeting Duration Calculator")
        self.root.geometry("1100x750")
        self.root.configure(bg="#f0f4f8")

        self.style = setup_styles()

        self.file_path = tk.StringVar()
        self.data = None
        self.result_data = None
        self.original_count = 0
        self.processed_count = 0

        main_frame = ttk.Frame(root, padding="15")
        main_frame.pack(fill="both", expand=True)

        # ----- Header / Title -----
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill="x", pady=(0, 10))

        title_label = ttk.Label(title_frame,
                                text="⏱️ Meeting Duration Calculator",
                                font=("Segoe UI", 18, "bold"),
                                foreground="#1a2b3c")
        title_label.pack(side="left")

        subtitle = ttk.Label(title_frame,
                             text="Calculate meeting durations from attendance logs",
                             font=("Segoe UI", 10),
                             foreground="#5a6b7c")
        subtitle.pack(side="left", padx=(15, 0))

        # ----- Description -----
        desc_text = (
            "📌 This app calculates meeting durations from attendance logs.\n"
            "   • Select a CSV or Excel file with columns: Time, User full name, Event name.\n"
            "   • The app pairs 'Meeting joined' and 'Meeting left' events for each user.\n"
            "   • Calculates duration for each meeting and sums total duration per user.\n"
            "   • Output includes: User, Joined, Left, Duration (HH:MM:SS).\n"
            "   • Download results as CSV or Excel."
        )
        desc_label = ttk.Label(main_frame, text=desc_text, font=("Segoe UI", 9),
                               foreground="#2a4b6c", background="#f0f4f8",
                               justify="left", wraplength=1000)
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

        process_btn = ttk.Button(file_frame, text="⚡ Calculate Durations", command=self.process_file,
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
                                 columns=("User", "Joined", "Left", "Duration"),
                                 show="headings",
                                 height=20)

        self.tree.heading("User", text="User Full Name")
        self.tree.heading("Joined", text="Joined Time")
        self.tree.heading("Left", text="Left Time")
        self.tree.heading("Duration", text="Duration (HH:MM:SS)")

        self.tree.column("User", width=300, anchor="w")
        self.tree.column("Joined", width=200, anchor="center")
        self.tree.column("Left", width=200, anchor="center")
        self.tree.column("Duration", width=150, anchor="center")

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

    def parse_time(self, time_str):
        """Parse time string to datetime object"""
        try:
            # Try parsing with 24-hour format (20/08/26, 15:54:08)
            dt = datetime.strptime(time_str.strip(), "%y/%m/%d, %H:%M:%S")
            return dt
        except:
            try:
                # Try parsing with 12-hour format (8/20/2026, 01:49:43 PM)
                dt = datetime.strptime(time_str.strip(), "%m/%d/%Y, %I:%M:%S %p")
                return dt
            except:
                try:
                    # Try parsing with 12-hour format without AM/PM
                    dt = datetime.strptime(time_str.strip(), "%m/%d/%Y, %I:%M:%S")
                    return dt
                except:
                    try:
                        # Try alternative format (20/08/2026, 15:54:08)
                        dt = datetime.strptime(time_str.strip(), "%d/%m/%Y, %H:%M:%S")
                        return dt
                    except:
                        return None

    def format_duration(self, td):
        """Format timedelta as HH:MM:SS"""
        total_seconds = int(td.total_seconds())
        hours, rem = divmod(total_seconds, 3600)
        minutes, seconds = divmod(rem, 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

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

            # Filter only Meeting joined and Meeting left events
            meeting_events = df[df["Event name"].str.contains("Meeting", case=False, na=False)]
            
            if len(meeting_events) == 0:
                messagebox.showwarning("Warning", "No meeting events found in the file.")
                return

            # Parse times
            meeting_events["Parsed_Time"] = meeting_events["Time"].apply(self.parse_time)
            
            # Remove rows with unparseable times
            meeting_events = meeting_events.dropna(subset=["Parsed_Time"])
            
            if len(meeting_events) == 0:
                messagebox.showerror("Error", "Could not parse time values. Please check the time format.")
                return

            # Sort by user and time
            meeting_events = meeting_events.sort_values(["User full name", "Parsed_Time"])

            # Group by user and process events
            results = []
            
            for user, group in meeting_events.groupby("User full name"):
                joined_times = []
                left_times = []
                
                # Process events in chronological order
                for _, row in group.iterrows():
                    if "joined" in row["Event name"].lower():
                        joined_times.append(row["Parsed_Time"])
                    elif "left" in row["Event name"].lower():
                        left_times.append(row["Parsed_Time"])
                
                # Pair joined and left times
                # If there are more joins than leaves, use the earliest join for each leave
                # If there are more leaves than joins, use the latest join for each leave
                i = 0
                j = 0
                while i < len(joined_times) and j < len(left_times):
                    # If join time is before leave time, it's a valid pair
                    if joined_times[i] < left_times[j]:
                        duration = left_times[j] - joined_times[i]
                        results.append({
                            "User": user,
                            "Joined": joined_times[i].strftime("%Y-%m-%d %H:%M:%S"),
                            "Left": left_times[j].strftime("%Y-%m-%d %H:%M:%S"),
                            "Duration": duration
                        })
                        i += 1
                        j += 1
                    else:
                        # If join is after leave, skip this leave (it might be from a previous session)
                        j += 1
                
                # If there are leftover joins without leaves, we can't calculate duration
                # If there are leftover leaves without joins, we can't calculate duration

            if not results:
                messagebox.showwarning("Warning", "Could not pair any meeting joined/left events.")
                return

            # Create DataFrame from results
            result_df = pd.DataFrame(results)
            
            # Calculate total duration per user
            total_duration_df = result_df.groupby("User").agg({
                "Joined": "first",  # First join time
                "Left": "last",     # Last left time
                "Duration": "sum"   # Sum of all durations
            }).reset_index()
            
            # Format duration as HH:MM:SS
            total_duration_df["Duration"] = total_duration_df["Duration"].apply(self.format_duration)
            
            # Reorder columns
            total_duration_df = total_duration_df[["User", "Joined", "Left", "Duration"]]
            
            self.processed_count = len(total_duration_df)
            self.result_data = total_duration_df

            # Populate treeview
            self.populate_tree(total_duration_df)

            # Update status
            self.status.config(
                text=f"✅ Processed: {self.original_count} rows → {self.processed_count} unique users with durations"
            )
            self.count_label.config(
                text=f"📊 Total events: {self.original_count} | Users with durations: {self.processed_count}"
            )
            
            messagebox.showinfo("Success",
                                f"Duration calculation complete!\n"
                                f"Total events processed: {self.original_count}\n"
                                f"Users with calculated durations: {self.processed_count}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to process file:\n{str(e)}")
            self.status.config(text="❌ Error")
            import traceback
            traceback.print_exc()

    def populate_tree(self, df):
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Insert data
        for _, row in df.iterrows():
            self.tree.insert("", "end", values=[row["User"], row["Joined"], row["Left"], row["Duration"]])

    def save_result(self, format_type):
        if self.result_data is None or self.result_data.empty:
            messagebox.showwarning("Warning", "No data to save. Process a file first.")
            return

        # Generate default filename
        base_name = "meeting_durations"
        if format_type == "csv":
            file_types = [("CSV files", "*.csv")]
            default_ext = ".csv"
        else:
            file_types = [("Excel files", "*.xlsx")]
            default_ext = ".xlsx"

        file_path = filedialog.asksaveasfilename(
            title=f"Save Meeting Durations as {format_type.upper()}",
            defaultextension=default_ext,
            filetypes=file_types,
            initialfile=f"{base_name}{default_ext}"
        )
        
        if file_path:
            try:
                if format_type == "csv":
                    self.result_data.to_csv(file_path, index=False)
                else:
                    self.result_data.to_excel(file_path, index=False)
                
                messagebox.showinfo("Success", f"File saved to:\n{file_path}")
                self.status.config(text=f"💾 Saved: {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not save file:\n{str(e)}")

    def clear_all(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.data = None
        self.result_data = None
        self.file_path.set("")
        self.original_count = 0
        self.processed_count = 0
        self.status.config(text="Cleared")
        self.count_label.config(text="")

if __name__ == "__main__":
    root = tk.Tk()
    app = AttendanceDurationCalculator(root)
    root.mainloop()