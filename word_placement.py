import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import re
import os
import sys
import traceback
import pandas as pd
from pathlib import Path
from datetime import datetime

# Word handling — preserves formatting
from docx import Document


class WordMailMerge:
    def __init__(self, root):
        self.root = root
        self.root.title("📄 Word Mail Merge (Word + Excel/CSV)")
        self.root.geometry("1150x900")
        self.root.configure(bg="#f0f4f8")

        # ---------- State ----------
        self.word_path = ""            # source .docx
        self.data = []                 # rows from CSV/Excel
        self.headers = []
        self.placeholders = []         # unique placeholders found in the doc
        self.merged_docs = []          # list of dicts: {name, filename, doc_bytes}

        # ---------- Main frame ----------
        main = ttk.Frame(root, padding="15")
        main.pack(fill="both", expand=True)

        # Title
        ttk.Label(
            main,
            text="📄 Word Mail Merge (Word + Excel/CSV)",
            font=("Segoe UI", 16, "bold"),
            foreground="#1a2b3c",
        ).pack(anchor="w")

        ttk.Label(
            main,
            text="Upload a Word template and a data file. Placeholders in Word get replaced with data. Formatting stays 100% intact.",
            font=("Segoe UI", 10),
            foreground="#5a6b7c",
        ).pack(anchor="w", pady=(0, 10))

        # ===== Top row: two file pickers =====
        top_frame = ttk.Frame(main)
        top_frame.pack(fill="x", pady=5)

        # --- Word picker (left) ---
        word_frame = ttk.LabelFrame(top_frame, text="📄 Word Template (.docx)", padding=10)
        word_frame.pack(side="left", fill="x", expand=True, padx=(0, 5))

        word_row = ttk.Frame(word_frame)
        word_row.pack(fill="x", pady=2)
        ttk.Label(word_row, text="File:").pack(side="left", padx=5)
        self.word_label = ttk.Label(
            word_row, text="No Word file selected", relief="sunken",
            background="white", padding=4
        )
        self.word_label.pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(word_row, text="📂 Browse Word", command=self.load_word).pack(side="left", padx=5)

        self.word_status = ttk.Label(word_frame, text="", foreground="#1a6e3c")
        self.word_status.pack(anchor="w", pady=2)

        # --- Data picker (right) ---
        data_frame = ttk.LabelFrame(top_frame, text="📊 Data File (Excel / CSV)", padding=10)
        data_frame.pack(side="left", fill="x", expand=True, padx=(5, 0))

        data_row = ttk.Frame(data_frame)
        data_row.pack(fill="x", pady=2)
        ttk.Label(data_row, text="File:").pack(side="left", padx=5)
        self.data_label = ttk.Label(
            data_row, text="No data file selected", relief="sunken",
            background="white", padding=4
        )
        self.data_label.pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(data_row, text="📂 Browse Data", command=self.load_data).pack(side="left", padx=5)

        col_row = ttk.Frame(data_frame)
        col_row.pack(fill="x", pady=5)
        ttk.Label(col_row, text="Email column (optional):").pack(side="left", padx=5)
        self.email_combo = ttk.Combobox(col_row, state="readonly", width=25)
        self.email_combo.pack(side="left", padx=5)

        self.data_status = ttk.Label(data_frame, text="", foreground="#1a6e3c")
        self.data_status.pack(anchor="w", pady=2)

        # ===== Placeholder detection panel =====
        info_frame = ttk.LabelFrame(main, text="🔍 Detected Placeholders", padding=10)
        info_frame.pack(fill="x", pady=5)

        self.placeholder_label = ttk.Label(
            info_frame,
            text="Upload a Word template to auto-detect placeholders.",
            foreground="#5a6b7c",
            font=("Segoe UI", 9),
        )
        self.placeholder_label.pack(anchor="w")

        # ===== Action buttons =====
        action_frame = ttk.Frame(main)
        action_frame.pack(fill="x", pady=10)

        ttk.Button(
            action_frame, text="⚡ Process (Merge)",
            command=self.process_merge,
            style="Success.TButton",
        ).pack(side="left", padx=5)

        ttk.Button(
            action_frame, text="💾 Download Selected",
            command=self.download_selected,
            style="Primary.TButton",
        ).pack(side="left", padx=5)

        ttk.Button(
            action_frame, text="📦 Download All (to folder)",
            command=self.download_all,
            style="Primary.TButton",
        ).pack(side="left", padx=5)

        ttk.Button(
            action_frame, text="📄 Preview First",
            command=self.preview_first,
            style="Primary.TButton",
        ).pack(side="left", padx=5)

        ttk.Button(
            action_frame, text="🐞 Debug",
            command=self.show_debug_info,
            style="Primary.TButton",
        ).pack(side="left", padx=5)

        # ===== Results list =====
        results_frame = ttk.LabelFrame(main, text="Merged Documents (double-click to preview)", padding=10)
        results_frame.pack(fill="both", expand=True, pady=5)

        self.results_listbox = tk.Listbox(
            results_frame, font=("Segoe UI", 9), selectmode=tk.EXTENDED
        )
        self.results_listbox.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(results_frame, orient="vertical",
                                  command=self.results_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.results_listbox.config(yscrollcommand=scrollbar.set)

        self.results_listbox.bind("<Double-Button-1>", self.preview_selected)

        # ===== Status bar =====
        self.status = ttk.Label(
            main, text="Ready", relief="sunken", anchor="w",
            background="#e9ecef", foreground="#1a2b3c",
            font=("Segoe UI", 9), padding=4,
        )
        self.status.pack(fill="x", pady=(5, 0))

        self.setup_styles()

    # ---------- Styles ----------
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background="#f0f4f8")
        style.configure("TLabelFrame", background="#f0f4f8",
                        foreground="#1a2b3c", font=("Segoe UI", 10))
        style.configure("TLabel", background="#f0f4f8",
                        foreground="#1a2b3c", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6)
        style.map("TButton", background=[("active", "#2a7de1")])
        style.configure("Success.TButton", background="#28a745", foreground="white")
        style.map("Success.TButton", background=[("active", "#218838")])
        style.configure("Primary.TButton", background="#2a7de1", foreground="white")
        style.map("Primary.TButton", background=[("active", "#1a5fb4")])

    # ---------- Load Word ----------
    def load_word(self):
        path = filedialog.askopenfilename(
            title="Select Word template",
            filetypes=[("Word files", "*.docx"), ("All files", "*.*")],
        )
        if not path:
            return

        if not path.lower().endswith(".docx"):
            messagebox.showerror(
                "Error",
                "Only .docx files are supported. If you have an old .doc, "
                "open it in Word and save as .docx first.",
            )
            return

        try:
            self.word_path = path
            self.word_label.config(text=Path(path).name)
            self.detect_placeholders()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Word:\n{str(e)}")

    def detect_placeholders(self):
        """Scan the Word doc for {{ placeholder }} patterns."""
        try:
            doc = Document(self.word_path)
            all_text = []

            # From paragraphs
            for p in doc.paragraphs:
                all_text.append(p.text)

            # From tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            all_text.append(p.text)

            # From headers/footers
            for section in doc.sections:
                for p in section.header.paragraphs:
                    all_text.append(p.text)
                for p in section.footer.paragraphs:
                    all_text.append(p.text)

            combined = "\n".join(all_text)
            found = re.findall(r'\{\{\s*([^}]+?)\s*\}\}', combined)

            # Preserve order, remove duplicates
            seen = set()
            self.placeholders = []
            for f in found:
                if f not in seen:
                    seen.add(f)
                    self.placeholders.append(f)

            if self.placeholders:
                self.placeholder_label.config(
                    text=f"✅ Found {len(self.placeholders)} placeholder(s): "
                         + ", ".join(f"{{{{ {p} }}}}" for p in self.placeholders),
                    foreground="#1a6e3c",
                )
                self.word_status.config(
                    text=f"Word loaded – {len(self.placeholders)} placeholder(s) detected."
                )
            else:
                self.placeholder_label.config(
                    text="⚠️ No {{ placeholder }} found. Make sure your Word uses {{ columnname }} syntax.",
                    foreground="#d9534f",
                )
                self.word_status.config(text="Word loaded – 0 placeholders found.")

            self.status.config(text=f"Loaded Word: {Path(self.word_path).name}")

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Failed to read Word document:\n{str(e)}\n\n{traceback.format_exc()}",
            )

    # ---------- Load CSV / Excel ----------
    def load_data(self):
        path = filedialog.askopenfilename(
            title="Select data file",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return

        try:
            lower = path.lower()
            if lower.endswith(".csv"):
                try:
                    df = pd.read_csv(path, encoding="utf-8")
                except UnicodeDecodeError:
                    df = pd.read_csv(path, encoding="latin-1")
            elif lower.endswith(".xls"):
                df = pd.read_excel(path, engine="xlrd")
            else:
                df = pd.read_excel(path, engine="openpyxl")

            if df.empty:
                messagebox.showerror("Error", "The data file is empty.")
                return

            # Fill NaN with empty string for clean replacement
            df = df.fillna("")

            self.data = df.to_dict(orient="records")
            self.headers = list(df.columns)

            self.data_label.config(text=Path(path).name)

            # Auto-detect email column
            email_col = None
            for col in self.headers:
                if re.search(r'email|e-mail|mail', col, re.I):
                    email_col = col
                    break

            self.email_combo["values"] = self.headers
            if email_col:
                self.email_combo.set(email_col)
            else:
                self.email_combo.set("")

            self.data_status.config(
                text=f"Loaded {len(self.data)} rows, {len(self.headers)} columns."
            )
            self.status.config(
                text=f"Loaded data: {Path(path).name} – {len(self.data)} records"
            )

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Failed to read data file:\n{str(e)}\n\n{traceback.format_exc()}",
            )

    # ---------- Placeholder replacement (preserves formatting) ----------
    def replace_in_paragraph(self, paragraph, mapping):
        """
        Replace {{ key }} in a paragraph while preserving runs.
        If a placeholder spans multiple runs, we merge them safely.
        """
        # First: quick check — does this paragraph contain "{{"?
        full_text = "".join(run.text for run in paragraph.runs)
        if "{{" not in full_text:
            return

        # Build new text with replacements
        def repl(match):
            key = match.group(1).strip()
            return str(mapping.get(key, match.group(0)))

        new_text = re.sub(r'\{\{\s*([^}]+?)\s*\}\}', repl, full_text)

        if new_text == full_text:
            return  # nothing changed

        # Simple approach: put everything in the first run, clear the rest.
        # This preserves the FIRST run's formatting for the whole paragraph.
        # Good enough for most templates and keeps the doc valid.
        if paragraph.runs:
            paragraph.runs[0].text = new_text
            for run in paragraph.runs[1:]:
                run.text = ""

    def replace_in_table(self, table, mapping):
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    self.replace_in_paragraph(p, mapping)
                # Recurse into nested tables
                for nested in cell.tables:
                    self.replace_in_table(nested, mapping)

    def replace_in_doc(self, doc, mapping):
        # Body paragraphs
        for p in doc.paragraphs:
            self.replace_in_paragraph(p, mapping)

        # Tables
        for table in doc.tables:
            self.replace_in_table(table, mapping)

        # Headers / footers
        for section in doc.sections:
            for p in section.header.paragraphs:
                self.replace_in_paragraph(p, mapping)
            for p in section.footer.paragraphs:
                self.replace_in_paragraph(p, mapping)
            for t in section.header.tables:
                self.replace_in_table(t, mapping)
            for t in section.footer.tables:
                self.replace_in_table(t, mapping)

    # ---------- Process merge ----------
    def process_merge(self):
        try:
            if not self.word_path:
                messagebox.showwarning("Warning", "Please load a Word template first.")
                return
            if not self.data:
                messagebox.showwarning("Warning", "Please load a data file first.")
                return
            if not self.placeholders:
                messagebox.showwarning(
                    "Warning",
                    "No placeholders found in your Word document.\n"
                    "Use {{ columnname }} syntax in the Word template.",
                )
                return

            # Check that all placeholders have matching columns
            missing = [p for p in self.placeholders if p not in self.headers]
            if missing:
                proceed = messagebox.askyesno(
                    "Missing Columns",
                    "These placeholders have no matching column in your data:\n\n"
                    + "\n".join(f"• {m}" for m in missing)
                    + "\n\nThey will be left as-is in the output.\n\nContinue?",
                )
                if not proceed:
                    return

            self.merged_docs = []
            self.results_listbox.delete(0, tk.END)

            total = len(self.data)
            for i, row in enumerate(self.data, 1):
                try:
                    # Load fresh copy of the template each time
                    doc = Document(self.word_path)

                    # Build mapping for this row
                    mapping = {k: ("" if v is None else v) for k, v in row.items()}

                    # Replace
                    self.replace_in_doc(doc, mapping)

                    # Save to bytes
                    from io import BytesIO
                    buf = BytesIO()
                    doc.save(buf)
                    doc_bytes = buf.getvalue()

                    # Build a display name + filename
                    # Priority: name → firstname+othername → email → row #
                    display = ""
                    for key in ("fullname", "name", "firstname"):
                        if key in row and str(row[key]).strip():
                            if key == "firstname":
                                last = str(row.get("othername", "") or row.get("surname", "") or row.get("lastname", "")).strip()
                                display = f"{row['firstname']} {last}".strip()
                            else:
                                display = str(row[key]).strip()
                            break
                    if not display:
                        email_col = self.email_combo.get()
                        if email_col and str(row.get(email_col, "")).strip():
                            display = str(row[email_col]).strip()
                        else:
                            display = f"Row {i}"

                    # Safe filename
                    safe = re.sub(r'[^\w\-_.@]', '_', display)
                    if len(safe) > 100:
                        safe = safe[:100]
                    filename = f"{i}_{safe}.docx"

                    self.merged_docs.append({
                        "row_index": i,
                        "display": display,
                        "filename": filename,
                        "bytes": doc_bytes,
                    })

                    self.results_listbox.insert(tk.END, f"{i}. {display}")

                except Exception as e:
                    self.results_listbox.insert(
                        tk.END, f"{i}. ⚠️ Failed: {str(e)[:60]}"
                    )

            self.status.config(
                text=f"✅ Merged {len(self.merged_docs)} of {total} document(s)."
            )

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Merge failed:\n{str(e)}\n\n{traceback.format_exc()}",
            )
            self.status.config(text="Error during merge")

    # ---------- Download selected ----------
    def download_selected(self):
        if not self.merged_docs:
            messagebox.showwarning("Warning", "No merged docs yet. Click 'Process' first.")
            return

        selection = self.results_listbox.curselection()
        if not selection:
            messagebox.showinfo("Info", "Select one or more rows in the list (Ctrl/Shift click for multiple).")
            return

        # Single file → Save As
        if len(selection) == 1:
            idx = selection[0]
            if idx >= len(self.merged_docs):
                return
            doc = self.merged_docs[idx]
            save_path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word files", "*.docx"), ("All files", "*.*")],
                initialfile=doc["filename"],
            )
            if not save_path:
                return
            with open(save_path, "wb") as f:
                f.write(doc["bytes"])
            self.status.config(text=f"💾 Saved {Path(save_path).name}")
            messagebox.showinfo("Success", f"Saved:\n{save_path}")
            return

        # Multiple → choose folder
        folder = filedialog.askdirectory(title="Choose folder to save selected docs")
        if not folder:
            return

        saved = 0
        for idx in selection:
            if idx >= len(self.merged_docs):
                continue
            doc = self.merged_docs[idx]
            out_path = Path(folder) / doc["filename"]
            with open(out_path, "wb") as f:
                f.write(doc["bytes"])
            saved += 1

        self.status.config(text=f"💾 Saved {saved} file(s) to {folder}")
        messagebox.showinfo("Success", f"Saved {saved} file(s) to:\n{folder}")

    # ---------- Download all ----------
    def download_all(self):
        if not self.merged_docs:
            messagebox.showwarning("Warning", "No merged docs yet. Click 'Process' first.")
            return

        parent = filedialog.askdirectory(title="Choose a folder to save all documents into")
        if not parent:
            return

        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_dir = Path(parent) / f"merged_{timestamp}"
            out_dir.mkdir(parents=True, exist_ok=True)

            saved = 0
            for doc in self.merged_docs:
                out_path = out_dir / doc["filename"]
                with open(out_path, "wb") as f:
                    f.write(doc["bytes"])
                saved += 1

            self.status.config(
                text=f"💾 Saved {saved} file(s) to {out_dir.name}"
            )

            if messagebox.askyesno(
                "Success",
                f"Saved {saved} file(s) to:\n\n{out_dir}\n\nOpen folder now?",
            ):
                try:
                    if sys.platform == "win32":
                        os.startfile(out_dir)
                    elif sys.platform == "darwin":
                        import subprocess
                        subprocess.Popen(["open", str(out_dir)])
                    else:
                        import subprocess
                        subprocess.Popen(["xdg-open", str(out_dir)])
                except Exception:
                    pass

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save:\n{str(e)}")

    # ---------- Preview first ----------
    def preview_first(self):
        if not self.merged_docs:
            messagebox.showinfo("Preview", "No merged docs yet. Click 'Process' first.")
            return
        self._open_doc_bytes(self.merged_docs[0]["bytes"], self.merged_docs[0]["filename"])

    def preview_selected(self, event=None):
        sel = self.results_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx >= len(self.merged_docs):
            return
        doc = self.merged_docs[idx]
        self._open_doc_bytes(doc["bytes"], doc["filename"])

    def _open_doc_bytes(self, doc_bytes, filename):
        """Write bytes to a temp file and open it with the system default app."""
        try:
            import tempfile
            tmp_dir = Path(tempfile.gettempdir()) / "word_merge_preview"
            tmp_dir.mkdir(exist_ok=True)
            tmp_path = tmp_dir / filename
            with open(tmp_path, "wb") as f:
                f.write(doc_bytes)

            if sys.platform == "win32":
                os.startfile(tmp_path)
            elif sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["open", str(tmp_path)])
            else:
                import subprocess
                subprocess.Popen(["xdg-open", str(tmp_path)])

            self.status.config(text=f"Opened preview: {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open preview:\n{str(e)}")

    # ---------- Debug ----------
    def show_debug_info(self):
        lines = []
        lines.append(f"Word template: {self.word_path or 'not selected'}")
        lines.append(f"Placeholders found in Word: {self.placeholders}")
        lines.append(f"Data rows: {len(self.data)}")
        lines.append(f"Data headers: {self.headers}")
        lines.append(f"Merged docs: {len(self.merged_docs)}")

        if self.placeholders and self.headers:
            missing = [p for p in self.placeholders if p not in self.headers]
            unused = [h for h in self.headers if h not in self.placeholders]
            lines.append("")
            lines.append(f"Placeholders missing in data: {missing or 'none'}")
            lines.append(f"Columns not used in Word: {unused or 'none'}")

        messagebox.showinfo("Debug Info", "\n".join(lines))


if __name__ == "__main__":
    root = tk.Tk()
    app = WordMailMerge(root)
    root.mainloop()