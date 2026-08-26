import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import re
import webbrowser
import urllib.parse
import pandas as pd
from pathlib import Path
import traceback
import sys

class WhatsAppLinkGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("📱 WhatsApp Link Generator (Excel)")
        self.root.geometry("1100x900")
        self.root.configure(bg="#f0f4f8")

        self.default_message = (
            "Hello {Name},\n\n"
            "Congratulations! You have been selected as one of the candidates sponsored to participate in the *IBS Youth Entrepreneurial Development Programme (IBS-YEDP) Cohort 2*.\n\n"
            "If you have not already joined the official WhatsApp group, please do so immediately. Important announcements, programme updates, orientation details, and instructions on accessing the programme before commencement will be communicated primarily through this group.\n\n"
            "*Join the WhatsApp group using the link below:*\n\n"
            "https://chat.whatsapp.com/Es6zmtqJlDtJMAwhNjLXj1?s=cl&p=a&ilr=1\n\n"
            "To avoid missing any important information, we encourage you to join as soon as possible.\n\n"
            "We look forward to welcoming you to the programme and wish you a rewarding learning experience.\n\n"
            "Warm regards,\n\n"
            "*School of Enterprise Management*\n"
            "*Ibadan Business School (IBS)*\n"
            "*IBS-YEDP Programme Team*"
        )

        self.links = []
        self.data = []
        self.headers = []
        self.phone_column = ""

        main = ttk.Frame(root, padding="15")
        main.pack(fill="both", expand=True)

        # Title
        ttk.Label(main, text="📱 WhatsApp Link Generator (Excel)",
                  font=("Segoe UI", 16, "bold"), foreground="#1a2b3c").pack(anchor="w")
        ttk.Label(main, text="Upload an Excel file, select the phone column, and personalize messages with placeholders like {Name}",
                  font=("Segoe UI", 10), foreground="#5a6b7c").pack(anchor="w", pady=(0, 10))

        # ----- Top row: Excel & debug -----
        top_frame = ttk.Frame(main)
        top_frame.pack(fill="x", pady=5)

        # Excel upload frame (left)
        excel_frame = ttk.LabelFrame(top_frame, text="Excel File", padding=10)
        excel_frame.pack(side="left", fill="x", expand=True, padx=(0, 5))

        file_row = ttk.Frame(excel_frame)
        file_row.pack(fill="x", pady=2)
        ttk.Label(file_row, text="File:").pack(side="left", padx=5)
        self.file_label = ttk.Label(file_row, text="No file selected", relief="sunken", background="white", padding=4)
        self.file_label.pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(file_row, text="📂 Browse Excel", command=self.load_excel).pack(side="left", padx=5)

        col_row = ttk.Frame(excel_frame)
        col_row.pack(fill="x", pady=5)
        ttk.Label(col_row, text="Phone number column:").pack(side="left", padx=5)
        self.phone_combo = ttk.Combobox(col_row, state="readonly", width=30)
        self.phone_combo.pack(side="left", padx=5)

        self.preview_label = ttk.Label(excel_frame, text="", foreground="#1a6e3c")
        self.preview_label.pack(anchor="w", pady=2)

        # Debug button (right)
        debug_frame = ttk.Frame(top_frame)
        debug_frame.pack(side="right", fill="y", padx=5)
        ttk.Button(debug_frame, text="🐞 Debug Info", command=self.show_debug_info,
                   style="Primary.TButton").pack(pady=2)
        ttk.Button(debug_frame, text="📄 Show First Link", command=self.show_first_link,
                   style="Primary.TButton").pack(pady=2)

        # ----- Message area -----
        msg_frame = ttk.LabelFrame(main, text="Message (use {ColumnName} placeholders)", padding=10)
        msg_frame.pack(fill="both", expand=True, pady=5)

        self.msg_text = scrolledtext.ScrolledText(msg_frame, height=10, font=("Segoe UI", 10))
        self.msg_text.pack(fill="both", expand=True)
        self.msg_text.insert("1.0", self.default_message)

        # ----- Action buttons -----
        action_frame = ttk.Frame(main)
        action_frame.pack(fill="x", pady=10)

        ttk.Button(action_frame, text="⚡ Generate Links", command=self.generate_links,
                   style="Success.TButton").pack(side="left", padx=5)
        ttk.Button(action_frame, text="🌐 Open All in Browser", command=self.open_all,
                   style="Primary.TButton").pack(side="left", padx=5)
        ttk.Button(action_frame, text="📋 Copy All Links", command=self.copy_links,
                   style="Primary.TButton").pack(side="left", padx=5)

        # ----- Generated links display -----
        links_frame = ttk.LabelFrame(main, text="Generated WhatsApp Links", padding=10)
        links_frame.pack(fill="both", expand=True, pady=5)

        self.links_listbox = tk.Listbox(links_frame, font=("Segoe UI", 9), selectmode=tk.EXTENDED)
        self.links_listbox.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(links_frame, orient="vertical", command=self.links_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.links_listbox.config(yscrollcommand=scrollbar.set)

        # Double‑click to open a single link
        self.links_listbox.bind("<Double-Button-1>", self.open_single_link)

        # ----- Status bar -----
        self.status = ttk.Label(main, text="Ready", relief="sunken", anchor="w",
                                background="#e9ecef", foreground="#1a2b3c", font=("Segoe UI", 9), padding=4)
        self.status.pack(fill="x", pady=(5, 0))

        self.setup_styles()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background="#f0f4f8")
        style.configure("TLabelFrame", background="#f0f4f8", foreground="#1a2b3c", font=("Segoe UI", 10))
        style.configure("TLabel", background="#f0f4f8", foreground="#1a2b3c", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6)
        style.map("TButton", background=[("active", "#2a7de1")])
        style.configure("Success.TButton", background="#28a745", foreground="white")
        style.map("Success.TButton", background=[("active", "#218838")])
        style.configure("Primary.TButton", background="#2a7de1", foreground="white")
        style.map("Primary.TButton", background=[("active", "#1a5fb4")])

    # ---------- Load Excel ----------
    def load_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not file_path:
            return
        try:
            if file_path.endswith('.xls'):
                df = pd.read_excel(file_path, engine='xlrd')
            else:
                df = pd.read_excel(file_path, engine='openpyxl')
            if df.empty:
                messagebox.showerror("Error", "The Excel file is empty.")
                return
            self.data = df.to_dict(orient='records')
            self.headers = list(df.columns)

            self.file_label.config(text=Path(file_path).name)

            # Auto-detect phone column
            phone_col = None
            for col in self.headers:
                if re.search(r'phone|number|mobile|cell', col, re.I):
                    phone_col = col
                    break
            self.phone_combo['values'] = self.headers
            if phone_col:
                self.phone_combo.set(phone_col)
                self.phone_column = phone_col
            else:
                self.phone_combo.set('')
                self.phone_column = ''

            self.preview_label.config(text=f"Loaded {len(self.data)} rows, {len(self.headers)} columns.")
            self.status.config(text=f"Loaded {Path(file_path).name} – {len(self.data)} records")
            self.links_listbox.delete(0, tk.END)
            self.links = []

        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file:\n{str(e)}\n\n{traceback.format_exc()}")
            self.status.config(text="Error loading file")

    # ---------- Generate links ----------
    def generate_links(self):
        try:
            if not self.data:
                messagebox.showwarning("Warning", "Please load an Excel file first.")
                return

            self.phone_column = self.phone_combo.get()
            if not self.phone_column:
                messagebox.showwarning("Warning", "Please select the phone number column.")
                return

            if self.phone_column not in self.headers:
                messagebox.showerror("Error", f"Column '{self.phone_column}' not found.")
                return

            message = self.msg_text.get("1.0", tk.END).strip()
            if not message:
                messagebox.showwarning("Warning", "Please enter a message.")
                return

            placeholder_pattern = re.compile(r'\{([^}]+)\}')
            placeholders = set(placeholder_pattern.findall(message))

            self.links = []
            skipped = 0
            for row in self.data:
                phone_raw = str(row.get(self.phone_column, '')).strip()
                if not phone_raw:
                    skipped += 1
                    continue
                phone_clean = re.sub(r'[^\d]', '', phone_raw)
                if not phone_clean:
                    skipped += 1
                    continue

                personalized_msg = message
                for placeholder in placeholders:
                    value = row.get(placeholder, '')
                    if value is None:
                        value = ''
                    personalized_msg = personalized_msg.replace(f'{{{placeholder}}}', str(value))

                encoded_msg = urllib.parse.quote(personalized_msg)
                link = f"https://wa.me/{phone_clean}?text={encoded_msg}"
                self.links.append(link)

            self.links_listbox.delete(0, tk.END)
            for link in self.links:
                self.links_listbox.insert(tk.END, link)

            status_text = f"✅ Generated {len(self.links)} link(s)."
            if skipped:
                status_text += f" Skipped {skipped} row(s) due to missing phone."
            self.status.config(text=status_text)

            # Show debug info if no links
            if len(self.links) == 0:
                self.show_debug_info()

        except Exception as e:
            error_msg = f"Error generating links:\n{str(e)}\n\nTraceback:\n{traceback.format_exc()}"
            messagebox.showerror("Error", error_msg)
            self.status.config(text="Error – see popup")

    # ---------- Debug: show first few rows ----------
    def show_debug_info(self):
        if not self.data:
            messagebox.showinfo("Debug", "No data loaded.")
            return

        phone_col = self.phone_combo.get() or "not selected"
        # Build a debug message
        lines = []
        lines.append(f"Total rows: {len(self.data)}")
        lines.append(f"Headers: {self.headers}")
        lines.append(f"Selected phone column: '{phone_col}'")
        lines.append("\nFirst 5 rows (raw data):")
        for i, row in enumerate(self.data[:5]):
            lines.append(f"Row {i+1}: {row}")

        # Also show the phone column values for the first 5
        if phone_col in self.headers:
            lines.append(f"\nPhone values (first 5, cleaned):")
            for i, row in enumerate(self.data[:5]):
                raw = str(row.get(phone_col, ''))
                clean = re.sub(r'[^\d]', '', raw)
                lines.append(f"  Row {i+1}: raw='{raw}' → clean='{clean}'")
        else:
            lines.append(f"\nPhone column '{phone_col}' not found in data headers.")

        messagebox.showinfo("Debug Info", "\n".join(lines))

    # ---------- Show first generated link ----------
    def show_first_link(self):
        if not self.links:
            messagebox.showinfo("First Link", "No links generated yet. Run 'Generate Links' first.")
            return
        first = self.links[0]
        # Show in a message box and also open it
        messagebox.showinfo("First Link (copy from here)", first)
        try:
            webbrowser.open_new_tab(first)
        except Exception as e:
            self.status.config(text=f"Could not open: {str(e)[:50]}")

    # ---------- Open all ----------
    def open_all(self):
        if not self.links:
            messagebox.showwarning("Warning", "No links generated yet. Click 'Generate Links' first.")
            return
        count = 0
        for link in self.links:
            try:
                webbrowser.open_new_tab(link)
                count += 1
            except Exception as e:
                self.status.config(text=f"Error opening: {str(e)[:50]}")
                return
        self.status.config(text=f"🌐 Opened {count} link(s) in browser tabs.")

    # ---------- Open single (double-click) ----------
    def open_single_link(self, event):
        selection = self.links_listbox.curselection()
        if not selection:
            return
        index = selection[0]
        if index < len(self.links):
            link = self.links[index]
            try:
                webbrowser.open_new_tab(link)
                self.status.config(text=f"Opened link #{index+1} in browser")
            except Exception as e:
                self.status.config(text=f"Error opening: {str(e)[:50]}")

    # ---------- Copy all ----------
    def copy_links(self):
        if not self.links:
            messagebox.showwarning("Warning", "No links to copy.")
            return
        text = "\n".join(self.links)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status.config(text=f"📋 Copied {len(self.links)} link(s) to clipboard.")


if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppLinkGenerator(root)
    root.mainloop()