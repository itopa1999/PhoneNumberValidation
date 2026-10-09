import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import re
import os
import sys
import traceback
import shutil
import subprocess
import html
import pandas as pd
from pathlib import Path
from datetime import datetime
from io import BytesIO
import tempfile

# Word handling — preserves formatting
from docx import Document

# SMTP imports
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.utils import formataddr


# Columns whose names suggest they hold URLs / attachments
LINK_COLUMN_PATTERN = re.compile(
    r'(document|brochure|attach|pdf|link|file|url|drive|form|sheet|folder|video|photo|image)',
    re.I,
)
# Basic URL detector
URL_PATTERN = re.compile(r'https?://[^\s<>"]+', re.I)


class WordMailMergeEmailSender:
    def __init__(self, root):
        self.root = root
        self.root.title("📄 Word Mail Merge → PDF → Email Sender")
        self.root.geometry("1250x820")
        self.root.minsize(1000, 600)
        self.root.configure(bg="#f0f4f8")

        # ---------- Default subject ----------
        self.default_subject = (
            "CALL FOR NOMINATIONS ON PROJECT MANAGEMENT PROGRAMME AND IBS BROCHURE"
        )

        # ---------- Default body ----------
        self.default_body = (
            "Dear Sir/Ma,\n\n"
            "Compliments of the day. I am sharing with you the call for nominations for our upcoming "
            "Project Management training programme for your kind consideration.\n\n"
            "We look forward to receiving nominations of officers from your organization to participate "
            "in the training programme. Please find the brochure below:\n\n"
            "📄 Brochure: {{ document }}\n\n"
            "Feel free to reach out with any enquiries.\n\n"
            "Kind regards,\n\n"
            "Olatipe Fasuyi\n\n"
            "Coordinator, School of Development Management"
        )

        # ---------- State ----------
        self.word_path = ""
        self.data = []
        self.headers = []
        self.placeholders = []
        self.merged_docs = []
        self.email_column = ""
        self.cc_column = ""
        self.recipients = []
        self.attached_pdf_path = ""
        self.link_columns = set()
        self.pdf_engine = None
        self.soffice_path = None

        self.smtp_config = {
            "server": "smtp.gmail.com",
            "port": 587,
            "sender_email": "",
            "sender_password": "",
            "sender_name": "IBS-YEDP Team",
        }

        # ============================================================
        # SCROLLABLE ROOT
        # ============================================================
        # Outer container fills the window
        outer = ttk.Frame(root)
        outer.pack(fill="both", expand=True)

        # Canvas + vertical scrollbar
        self.canvas = tk.Canvas(outer, bg="#f0f4f8", highlightthickness=0)
        vscroll = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vscroll.set)

        vscroll.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        # Inner frame that actually holds widgets
        self.scroll_frame = ttk.Frame(self.canvas, padding="15")
        self.canvas_window = self.canvas.create_window(
            (0, 0), window=self.scroll_frame, anchor="nw"
        )

        # Keep inner frame as wide as the canvas
        def _on_canvas_configure(event):
            self.canvas.itemconfig(self.canvas_window, width=event.width)
        self.canvas.bind("<Configure>", _on_canvas_configure)

        # Update scrollregion when inner frame changes
        def _on_frame_configure(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.scroll_frame.bind("<Configure>", _on_frame_configure)

        # Mouse wheel scrolling — works on Windows, Mac, Linux
        def _on_mousewheel(event):
            delta = 0
            if event.num == 4:
                delta = -1
            elif event.num == 5:
                delta = 1
            else:
                delta = int(-1 * (event.delta / 120)) if event.delta else 0
                if delta == 0 and event.delta:
                    delta = -1 if event.delta > 0 else 1
            self.canvas.yview_scroll(delta, "units")

        def _bind_mousewheel(widget):
            widget.bind_all("<MouseWheel>", _on_mousewheel, add="+")
            widget.bind_all("<Button-4>", _on_mousewheel, add="+")
            widget.bind_all("<Button-5>", _on_mousewheel, add="+")

        # Only scroll when the mouse is over the canvas area
        self.canvas.bind("<Enter>", lambda e: _bind_mousewheel(self.canvas))
        self.canvas.bind("<Leave>", lambda e: (
            self.canvas.unbind_all("<MouseWheel>"),
            self.canvas.unbind_all("<Button-4>"),
            self.canvas.unbind_all("<Button-5>"),
        ))

        # Now use scroll_frame as our "main"
        main = self.scroll_frame

        # ============================================================
        # TITLE
        # ============================================================
        ttk.Label(
            main,
            text="📄 Word Mail Merge → PDF → Email Sender",
            font=("Segoe UI", 16, "bold"),
            foreground="#1a2b3c",
        ).pack(anchor="w")

        ttk.Label(
            main,
            text=(
                "Upload a Word template + data file. Placeholders get replaced, each doc is converted "
                "to PDF, and emails are sent with the PDF attached. Link columns (document, brochure, "
                "video, etc.) are auto-linked in the email body."
            ),
            font=("Segoe UI", 10),
            foreground="#5a6b7c",
            wraplength=1150,
            justify="left",
        ).pack(anchor="w", pady=(0, 10))

        # ============================================================
        # TOP ROW: Word + Data pickers (side by side)
        # ============================================================
        top_frame = ttk.Frame(main)
        top_frame.pack(fill="x", pady=5)

        # --- Word picker ---
        word_frame = ttk.LabelFrame(top_frame, text="📄 Word Template (.docx)", padding=10)
        word_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        word_row = ttk.Frame(word_frame)
        word_row.pack(fill="x", pady=2)
        ttk.Label(word_row, text="File:").pack(side="left", padx=5)
        self.word_label = ttk.Label(
            word_row, text="No Word file selected", relief="sunken",
            background="white", padding=4
        )
        self.word_label.pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(word_row, text="📂 Browse Word", command=self.load_word).pack(side="left", padx=5)

        self.word_status = ttk.Label(word_frame, text="", foreground="#1a6e3c",
                                     font=("Segoe UI", 9))
        self.word_status.pack(anchor="w", pady=2)

        # --- Data picker ---
        data_frame = ttk.LabelFrame(top_frame, text="📊 Data File (Excel / CSV)", padding=10)
        data_frame.pack(side="left", fill="both", expand=True, padx=(5, 0))

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
        ttk.Label(col_row, text="Email column:").pack(side="left", padx=5)
        self.email_combo = ttk.Combobox(col_row, state="readonly", width=18)
        self.email_combo.pack(side="left", padx=5)
        ttk.Label(col_row, text="CC column:").pack(side="left", padx=5)
        self.cc_combo = ttk.Combobox(col_row, state="readonly", width=18)
        self.cc_combo.pack(side="left", padx=5)

        self.data_status = ttk.Label(data_frame, text="", foreground="#1a6e3c",
                                     font=("Segoe UI", 9))
        self.data_status.pack(anchor="w", pady=2)

        # ============================================================
        # PLACEHOLDERS + LINK COLUMNS + PDF ENGINE (side by side, compact)
        # ============================================================
        info_row = ttk.Frame(main)
        info_row.pack(fill="x", pady=5)

        # Placeholders
        ph_frame = ttk.LabelFrame(info_row, text="🔍 Detected Placeholders", padding=8)
        ph_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        self.placeholder_label = ttk.Label(
            ph_frame,
            text="Upload a Word template to auto-detect placeholders.",
            foreground="#5a6b7c",
            font=("Segoe UI", 9),
            wraplength=350,
            justify="left",
        )
        self.placeholder_label.pack(anchor="w")

        # Link columns
        lk_frame = ttk.LabelFrame(info_row, text="🔗 Auto-linked Columns", padding=8)
        lk_frame.pack(side="left", fill="both", expand=True, padx=5)

        self.link_status = ttk.Label(
            lk_frame,
            text="Load a data file to detect link columns.",
            foreground="#5a6b7c",
            font=("Segoe UI", 9),
            wraplength=350,
            justify="left",
        )
        self.link_status.pack(anchor="w")

        # PDF engine
        conv_frame = ttk.LabelFrame(info_row, text="🖨️ PDF Engine", padding=8)
        conv_frame.pack(side="left", fill="both", expand=True, padx=(5, 0))

        self.conv_status = ttk.Label(
            conv_frame,
            text="Checking conversion availability...",
            foreground="#5a6b7c",
            font=("Segoe UI", 9),
            wraplength=350,
            justify="left",
        )
        self.conv_status.pack(anchor="w")
        self.check_pdf_engine()

        # ============================================================
        # SUBJECT
        # ============================================================
        subj_frame = ttk.LabelFrame(main, text="Email Subject (use {{ ColumnName }} placeholders)", padding=8)
        subj_frame.pack(fill="x", pady=5)

        self.subject_entry = ttk.Entry(subj_frame, font=("Segoe UI", 10))
        self.subject_entry.pack(fill="x")
        self.subject_entry.insert(0, self.default_subject)

        # ============================================================
        # BODY
        # ============================================================
        msg_frame = ttk.LabelFrame(
            main,
            text="Email Body (use {{ ColumnName }} placeholders, HTML <strong>...</strong> allowed)",
            padding=8,
        )
        msg_frame.pack(fill="x", pady=5)

        self.msg_text = scrolledtext.ScrolledText(
            msg_frame, height=8, font=("Segoe UI", 10), wrap="word"
        )
        self.msg_text.pack(fill="both", expand=True)
        self.msg_text.insert("1.0", self.default_body)

        # ============================================================
        # FALLBACK PDF
        # ============================================================
        pdf_frame = ttk.LabelFrame(main, text="📎 Fallback PDF (used only when conversion fails for a row)", padding=8)
        pdf_frame.pack(fill="x", pady=5)

        pdf_row = ttk.Frame(pdf_frame)
        pdf_row.pack(fill="x")

        ttk.Label(pdf_row, text="PDF File:").pack(side="left", padx=5)
        self.pdf_label = ttk.Label(
            pdf_row, text="No PDF attached", relief="sunken",
            background="white", padding=4
        )
        self.pdf_label.pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(pdf_row, text="📎 Attach PDF", command=self.attach_pdf).pack(side="left", padx=5)
        ttk.Button(pdf_row, text="✖ Clear", command=self.clear_pdf).pack(side="left", padx=5)

        # ============================================================
        # ACTION BUTTONS (all on one visible row)
        # ============================================================
        action_frame = ttk.LabelFrame(main, text="⚙️ Actions", padding=10)
        action_frame.pack(fill="x", pady=10)

        # Row 1: primary actions
        row1 = ttk.Frame(action_frame)
        row1.pack(fill="x", pady=3)

        ttk.Button(
            row1, text="⚡ Process (Merge → PDF → Build Emails)",
            command=self.process_all,
            style="Success.TButton",
        ).pack(side="left", padx=4)

        ttk.Button(
            row1, text="🚀 Send via SMTP",
            command=self.open_smtp_window,
            style="Success.TButton",
        ).pack(side="left", padx=4)

        # Row 2: secondary actions
        row2 = ttk.Frame(action_frame)
        row2.pack(fill="x", pady=3)

        ttk.Button(
            row2, text="💾 Download PDFs",
            command=self.download_all,
            style="Primary.TButton",
        ).pack(side="left", padx=4)

        ttk.Button(
            row2, text="📄 Preview First PDF",
            command=self.preview_first,
            style="Primary.TButton",
        ).pack(side="left", padx=4)

        ttk.Button(
            row2, text="📧 Preview First Email",
            command=self.preview_first_email,
            style="Primary.TButton",
        ).pack(side="left", padx=4)

        ttk.Button(
            row2, text="🐞 Debug",
            command=self.show_debug_info,
            style="Primary.TButton",
        ).pack(side="left", padx=4)

        # ============================================================
        # RESULTS LIST
        # ============================================================
        results_frame = ttk.LabelFrame(
            main,
            text="Merged Documents & Recipients (double-click to preview PDF)",
            padding=8,
        )
        results_frame.pack(fill="both", expand=True, pady=5)

        list_container = ttk.Frame(results_frame)
        list_container.pack(fill="both", expand=True)

        self.results_listbox = tk.Listbox(
            list_container, font=("Segoe UI", 9),
            selectmode=tk.EXTENDED, height=8,
        )
        self.results_listbox.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(list_container, orient="vertical",
                                  command=self.results_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.results_listbox.config(yscrollcommand=scrollbar.set)

        self.results_listbox.bind("<Double-Button-1>", self.preview_selected)

        # ============================================================
        # STATUS BAR (at the bottom of the scrollable area)
        # ============================================================
        self.status = ttk.Label(
            main, text="Ready", relief="sunken", anchor="w",
            background="#e9ecef", foreground="#1a2b3c",
            font=("Segoe UI", 9), padding=4,
        )
        self.status.pack(fill="x", pady=(5, 0))

        self.setup_styles()

        # Recompute scrollregion after layout settles
        self.root.after(200, lambda: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")
        ))

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

    # ================= LINK HELPERS =================
    def is_link_column(self, name: str) -> bool:
        if not name:
            return False
        return bool(LINK_COLUMN_PATTERN.search(str(name)))

    def normalize_url(self, url: str) -> str:
        if not url:
            return url
        url = url.strip()
        m = re.match(r'https?://drive\.google\.com/file/d/([^/]+)/', url)
        if m:
            return f"https://drive.google.com/file/d/{m.group(1)}/preview"
        m = re.match(r'https?://docs\.google\.com/([^/]+)/d/([^/]+)/', url)
        if m:
            return f"https://docs.google.com/{m.group(1)}/d/{m.group(2)}/preview"
        return url

    def make_html_link(self, url: str, label: str = None) -> str:
        if not url:
            return ""
        clean = self.normalize_url(url)
        safe_href = html.escape(clean, quote=True)
        safe_label = html.escape(label or clean)
        return (
            f'<a href="{safe_href}" target="_blank" '
            f'style="color:#2a7de1;text-decoration:underline;">{safe_label}</a>'
        )

    def render_value_for_body(self, key: str, value) -> str:
        if value is None:
            return ""
        sval = str(value).strip()
        if not sval:
            return ""

        looks_like_url = bool(URL_PATTERN.search(sval))
        if looks_like_url:
            if self.is_link_column(key):
                label = "Click here to view"
            else:
                label = sval
            return self.make_html_link(sval, label)

        return html.escape(sval)

    # ================= PDF ENGINE =================
    def check_pdf_engine(self):
        self.pdf_engine = None
        self.soffice_path = None

        try:
            import docx2pdf  # noqa
            self.pdf_engine = "docx2pdf"
            self.conv_status.config(
                text="✅ Using 'docx2pdf' (requires Microsoft Word on Windows/Mac).",
                foreground="#1a6e3c",
            )
            return
        except ImportError:
            pass

        candidates = [
            shutil.which("soffice"),
            shutil.which("libreoffice"),
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for c in candidates:
            if c and Path(c).exists():
                self.soffice_path = c
                self.pdf_engine = "libreoffice"
                self.conv_status.config(
                    text=f"✅ Using LibreOffice: {Path(c).name}",
                    foreground="#1a6e3c",
                )
                return

        self.conv_status.config(
            text="⚠️ No PDF engine found. Install 'docx2pdf' (pip install docx2pdf) "
                 "OR LibreOffice.",
            foreground="#d9534f",
        )

    def convert_docx_bytes_to_pdf_bytes(self, docx_bytes, base_filename):
        if not self.pdf_engine:
            return None, "No PDF engine available."

        tmp_dir = Path(tempfile.gettempdir()) / "word_merge_pdf"
        tmp_dir.mkdir(exist_ok=True)

        docx_path = tmp_dir / f"{base_filename}.docx"
        pdf_path = tmp_dir / f"{base_filename}.pdf"

        try:
            with open(docx_path, "wb") as f:
                f.write(docx_bytes)

            if pdf_path.exists():
                try:
                    pdf_path.unlink()
                except Exception:
                    pass

            if self.pdf_engine == "docx2pdf":
                from docx2pdf import convert as docx2pdf_convert
                docx2pdf_convert(str(docx_path), str(pdf_path))
                if not pdf_path.exists():
                    return None, "docx2pdf did not produce output (is Word installed?)"

            elif self.pdf_engine == "libreoffice":
                result = subprocess.run(
                    [
                        self.soffice_path,
                        "--headless",
                        "--norestore",
                        "--convert-to", "pdf",
                        "--outdir", str(tmp_dir),
                        str(docx_path),
                    ],
                    capture_output=True,
                    text=True,
                    timeout=120,
                )
                if not pdf_path.exists():
                    matches = list(tmp_dir.glob(f"{base_filename}*.pdf"))
                    if not matches:
                        return None, f"LibreOffice failed: {result.stderr or result.stdout}"
                    pdf_path = matches[0]

            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()

            return pdf_bytes, None

        except Exception as e:
            return None, str(e)
        finally:
            try:
                if docx_path.exists():
                    docx_path.unlink()
            except Exception:
                pass

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
        try:
            doc = Document(self.word_path)
            all_text = []
            for p in doc.paragraphs:
                all_text.append(p.text)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            all_text.append(p.text)
            for section in doc.sections:
                for p in section.header.paragraphs:
                    all_text.append(p.text)
                for p in section.footer.paragraphs:
                    all_text.append(p.text)

            combined = "\n".join(all_text)
            found = re.findall(r'\{\{\s*([^}]+?)\s*\}\}', combined)

            seen = set()
            self.placeholders = []
            for f in found:
                if f not in seen:
                    seen.add(f)
                    self.placeholders.append(f)

            if self.placeholders:
                self.placeholder_label.config(
                    text=f"✅ Found {len(self.placeholders)}: "
                         + ", ".join(f"{{{{ {p} }}}}" for p in self.placeholders),
                    foreground="#1a6e3c",
                )
                self.word_status.config(
                    text=f"Word loaded – {len(self.placeholders)} placeholder(s)."
                )
            else:
                self.placeholder_label.config(
                    text="⚠️ No {{ placeholder }} found. Use {{ columnname }} syntax.",
                    foreground="#d9534f",
                )
                self.word_status.config(text="Word loaded – 0 placeholders.")

            self.status.config(text=f"Loaded Word: {Path(self.word_path).name}")
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

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

            df = df.fillna("")
            self.data = df.to_dict(orient="records")
            self.headers = list(df.columns)
            self.data_label.config(text=Path(path).name)

            self.link_columns = set()
            for col in self.headers:
                if self.is_link_column(col):
                    self.link_columns.add(col)
                    continue
                for r in self.data[:5]:
                    v = str(r.get(col, ""))
                    if URL_PATTERN.search(v):
                        self.link_columns.add(col)
                        break

            if self.link_columns:
                self.link_status.config(
                    text="🔗 Linked: "
                         + ", ".join(f"{{{{ {c} }}}}" for c in sorted(self.link_columns)),
                    foreground="#1a6e3c",
                )
            else:
                self.link_status.config(
                    text="ℹ️ No link columns detected (document, brochure, link, file, video, etc.).",
                    foreground="#5a6b7c",
                )

            email_col = None
            for col in self.headers:
                if re.search(r'email|e-mail|mail', col, re.I):
                    email_col = col
                    break

            cc_col = None
            for col in self.headers:
                if re.search(r'^cc$|cc_|_cc|carbon', col, re.I):
                    cc_col = col
                    break

            self.email_combo["values"] = self.headers
            if email_col:
                self.email_combo.set(email_col)
                self.email_column = email_col
            else:
                self.email_combo.set("")
                self.email_column = ""

            self.cc_combo["values"] = ["(none)"] + self.headers
            if cc_col:
                self.cc_combo.set(cc_col)
                self.cc_column = cc_col
            else:
                self.cc_combo.set("(none)")
                self.cc_column = ""

            self.data_status.config(
                text=f"Loaded {len(self.data)} rows, {len(self.headers)} columns."
            )
            self.status.config(
                text=f"Loaded data: {Path(path).name} – {len(self.data)} records"
            )
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Failed to read data file:\n{str(e)}\n\n{traceback.format_exc()}",
            )

    # ---------- Attach fallback PDF ----------
    def attach_pdf(self):
        path = filedialog.askopenfilename(
            title="Select a PDF to attach as fallback",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not path:
            return
        self.attached_pdf_path = path
        size_kb = os.path.getsize(path) / 1024
        self.pdf_label.config(text=f"{Path(path).name}  ({size_kb:.1f} KB)")
        self.status.config(text=f"📎 Fallback PDF attached: {Path(path).name}")

    def clear_pdf(self):
        self.attached_pdf_path = ""
        self.pdf_label.config(text="No PDF attached")
        self.status.config(text="PDF attachment cleared.")

    # ---------- Placeholder replacement (Word doc) ----------
    def replace_in_paragraph(self, paragraph, mapping):
        full_text = "".join(run.text for run in paragraph.runs)
        if "{{" not in full_text:
            return

        def repl(match):
            key = match.group(1).strip()
            return str(mapping.get(key, match.group(0)))

        new_text = re.sub(r'\{\{\s*([^}]+?)\s*\}\}', repl, full_text)

        if new_text == full_text:
            return

        if paragraph.runs:
            paragraph.runs[0].text = new_text
            for run in paragraph.runs[1:]:
                run.text = ""

    def replace_in_table(self, table, mapping):
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    self.replace_in_paragraph(p, mapping)
                for nested in cell.tables:
                    self.replace_in_table(nested, mapping)

    def replace_in_doc(self, doc, mapping):
        for p in doc.paragraphs:
            self.replace_in_paragraph(p, mapping)
        for table in doc.tables:
            self.replace_in_table(table, mapping)
        for section in doc.sections:
            for p in section.header.paragraphs:
                self.replace_in_paragraph(p, mapping)
            for p in section.footer.paragraphs:
                self.replace_in_paragraph(p, mapping)
            for t in section.header.tables:
                self.replace_in_table(t, mapping)
            for t in section.footer.tables:
                self.replace_in_table(t, mapping)

    # ---------- Process all ----------
    def process_all(self):
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

            if not self.pdf_engine:
                if not messagebox.askyesno(
                    "No PDF Engine",
                    "No PDF conversion engine is available.\n\n"
                    "Install 'docx2pdf' (pip install docx2pdf) or LibreOffice.\n\n"
                    "Continue anyway? (PDF conversion will fail; fallback PDF will be used if provided.)"
                ):
                    return

            self.email_column = self.email_combo.get()
            if not self.email_column:
                messagebox.showwarning("Warning", "Please select the email column.")
                return

            self.cc_column = self.cc_combo.get()
            if self.cc_column == "(none)":
                self.cc_column = ""

            subject_template = self.subject_entry.get().strip()
            body_template = self.msg_text.get("1.0", tk.END).strip()

            if not subject_template:
                messagebox.showwarning("Warning", "Please enter an email subject.")
                return
            if not body_template:
                messagebox.showwarning("Warning", "Please enter an email body.")
                return

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
            self.recipients = []
            self.results_listbox.delete(0, tk.END)

            placeholder_pattern = re.compile(r'\{\{\s*([^}]+?)\s*\}\}')
            total = len(self.data)
            skipped = 0
            pdf_failures = 0

            self.status.config(text=f"Processing {total} row(s)...")
            self.root.update()

            for i, row in enumerate(self.data, 1):
                try:
                    mapping = {k: ("" if v is None else v) for k, v in row.items()}
                    doc = Document(self.word_path)
                    self.replace_in_doc(doc, mapping)

                    buf = BytesIO()
                    doc.save(buf)
                    docx_bytes = buf.getvalue()

                    display = ""
                    for key in ("fullname", "name", "firstname", "title"):
                        if key in row and str(row[key]).strip():
                            if key == "firstname":
                                last = str(
                                    row.get("othername", "")
                                    or row.get("surname", "")
                                    or row.get("lastname", "")
                                ).strip()
                                display = f"{row['firstname']} {last}".strip()
                            else:
                                display = str(row[key]).strip()
                            break
                    if not display:
                        email_val = str(row.get(self.email_column, "")).strip()
                        display = email_val if email_val else f"Row {i}"

                    safe = re.sub(r'[^\w\-_.@]', '_', display)
                    if len(safe) > 80:
                        safe = safe[:80]

                    docx_filename = f"{i}_{safe}.docx"
                    pdf_filename = f"{i}_{safe}.pdf"

                    pdf_bytes = None
                    pdf_err = None
                    if self.pdf_engine:
                        base_unique = f"{i}_{safe}_{datetime.now().strftime('%H%M%S%f')}"
                        pdf_bytes, pdf_err = self.convert_docx_bytes_to_pdf_bytes(
                            docx_bytes, base_unique
                        )
                        if pdf_err:
                            pdf_failures += 1

                    email_raw = str(row.get(self.email_column, "")).strip()
                    if not email_raw or "@" not in email_raw:
                        skipped += 1
                        self.results_listbox.insert(
                            tk.END,
                            f"{i}. ⚠️ Skipped (no valid email): {display}"
                        )
                        self.merged_docs.append({
                            "row_index": i,
                            "display": display,
                            "docx_filename": docx_filename,
                            "pdf_filename": pdf_filename,
                            "docx_bytes": docx_bytes,
                            "pdf_bytes": pdf_bytes,
                        })
                        continue

                    cc_raw = ""
                    if self.cc_column:
                        cc_raw = str(row.get(self.cc_column, "")).strip()

                    def replacer_plain(match):
                        key = match.group(1).strip()
                        value = row.get(key, "")
                        if value is None or (isinstance(value, float) and pd.isna(value)):
                            value = ""
                        return str(value)

                    def replacer_body(match):
                        key = match.group(1).strip()
                        value = row.get(key, "")
                        if value is None or (isinstance(value, float) and pd.isna(value)):
                            value = ""
                        return self.render_value_for_body(key, value)

                    personalized_subject = placeholder_pattern.sub(replacer_plain, subject_template)
                    personalized_body = placeholder_pattern.sub(replacer_body, body_template)

                    self.merged_docs.append({
                        "row_index": i,
                        "display": display,
                        "docx_filename": docx_filename,
                        "pdf_filename": pdf_filename,
                        "docx_bytes": docx_bytes,
                        "pdf_bytes": pdf_bytes,
                    })

                    self.recipients.append({
                        "email": email_raw,
                        "cc": cc_raw,
                        "name": display,
                        "subject": personalized_subject,
                        "body": personalized_body,
                        "docx_bytes": docx_bytes,
                        "docx_filename": docx_filename,
                        "pdf_bytes": pdf_bytes,
                        "pdf_filename": pdf_filename,
                    })

                    pdf_tag = "📎 PDF" if pdf_bytes else ("⚠️ No PDF" if self.pdf_engine else "⚠️ No engine")
                    cc_display = f" | CC: {cc_raw}" if cc_raw else ""
                    self.results_listbox.insert(
                        tk.END,
                        f"{i}. {display} <{email_raw}>{cc_display}  [{pdf_tag}]"
                    )

                except Exception as e:
                    self.results_listbox.insert(
                        tk.END, f"{i}. ⚠️ Failed: {str(e)[:60]}"
                    )

                self.status.config(text=f"Processing {i}/{total}...")
                self.root.update()

            status_text = (
                f"✅ Merged {len(self.merged_docs)} doc(s) → PDF, built {len(self.recipients)} email(s)."
            )
            if skipped:
                status_text += f" Skipped {skipped} row(s) with missing/invalid email."
            if pdf_failures:
                status_text += f" ⚠️ {pdf_failures} PDF conversion(s) failed."
            self.status.config(text=status_text)

            # Update scrollregion after results
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Processing failed:\n{str(e)}\n\n{traceback.format_exc()}",
            )
            self.status.config(text="Error during processing")

    # ---------- Download PDFs ----------
    def download_all(self):
        if not self.merged_docs:
            messagebox.showwarning("Warning", "No merged docs yet. Click 'Process' first.")
            return

        parent = filedialog.askdirectory(title="Choose a folder to save all PDFs into")
        if not parent:
            return

        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_dir = Path(parent) / f"merged_pdfs_{timestamp}"
            out_dir.mkdir(parents=True, exist_ok=True)

            saved_pdf = 0
            saved_docx = 0
            for doc in self.merged_docs:
                if doc.get("pdf_bytes"):
                    with open(out_dir / doc["pdf_filename"], "wb") as f:
                        f.write(doc["pdf_bytes"])
                    saved_pdf += 1
                if doc.get("docx_bytes"):
                    with open(out_dir / doc["docx_filename"], "wb") as f:
                        f.write(doc["docx_bytes"])
                    saved_docx += 1

            self.status.config(
                text=f"💾 Saved {saved_pdf} PDF(s) and {saved_docx} DOCX(s) to {out_dir.name}"
            )

            if messagebox.askyesno(
                "Success",
                f"Saved {saved_pdf} PDF(s) + {saved_docx} DOCX(s) to:\n\n{out_dir}\n\nOpen folder now?",
            ):
                try:
                    if sys.platform == "win32":
                        os.startfile(out_dir)
                    elif sys.platform == "darwin":
                        subprocess.Popen(["open", str(out_dir)])
                    else:
                        subprocess.Popen(["xdg-open", str(out_dir)])
                except Exception:
                    pass

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save:\n{str(e)}")

    # ---------- Preview ----------
    def preview_first(self):
        if not self.merged_docs:
            messagebox.showinfo("Preview", "No merged docs yet. Click 'Process' first.")
            return
        d = self.merged_docs[0]
        if d.get("pdf_bytes"):
            self._open_bytes(d["pdf_bytes"], d["pdf_filename"])
        elif d.get("docx_bytes"):
            self._open_bytes(d["docx_bytes"], d["docx_filename"])
        else:
            messagebox.showinfo("Preview", "No preview available for this row.")

    def preview_selected(self, event=None):
        sel = self.results_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx >= len(self.merged_docs):
            return
        d = self.merged_docs[idx]
        if d.get("pdf_bytes"):
            self._open_bytes(d["pdf_bytes"], d["pdf_filename"])
        elif d.get("docx_bytes"):
            self._open_bytes(d["docx_bytes"], d["docx_filename"])
        else:
            messagebox.showinfo("Preview", "No preview available for this row.")

    def _open_bytes(self, data_bytes, filename):
        try:
            tmp_dir = Path(tempfile.gettempdir()) / "word_merge_preview"
            tmp_dir.mkdir(exist_ok=True)
            tmp_path = tmp_dir / filename
            with open(tmp_path, "wb") as f:
                f.write(data_bytes)

            if sys.platform == "win32":
                os.startfile(tmp_path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(tmp_path)])
            else:
                subprocess.Popen(["xdg-open", str(tmp_path)])

            self.status.config(text=f"Opened preview: {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open preview:\n{str(e)}")

    def preview_first_email(self):
        if not self.recipients:
            messagebox.showinfo("Preview", "No processed emails yet. Click 'Process' first.")
            return
        r = self.recipients[0]
        cc_line = f"\nCC: {r['cc']}" if r.get("cc") else ""
        attachment = r.get("pdf_filename") if r.get("pdf_bytes") else r.get("docx_filename")

        preview = (
            f"To: {r['email']}{cc_line}\n"
            f"Name: {r['name']}\n"
            f"Attachment: {attachment}\n\n"
            f"Subject:\n{r['subject']}\n\n"
            f"{'-'*60}\n\n"
            f"{r['body']}\n\n"
            f"{'-'*60}\n"
            f"(HTML body — links should appear as clickable links in the email.)"
        )
        messagebox.showinfo("First Email Preview", preview)

    # ---------- SMTP Window ----------
    def open_smtp_window(self):
        if not self.recipients:
            messagebox.showwarning("Warning", "Process first (click 'Process').")
            return

        win = tk.Toplevel(self.root)
        win.title("🚀 Send via SMTP")
        win.geometry("560x640")
        win.minsize(500, 500)
        win.configure(bg="#f0f4f8")
        win.transient(self.root)
        win.grab_set()

        ttk.Label(win, text="SMTP Settings", font=("Segoe UI", 13, "bold"),
                  foreground="#1a2b3c", background="#f0f4f8").pack(pady=(15, 10))

        form = ttk.Frame(win, padding=10)
        form.pack(fill="x", padx=10)

        fields = {}
        rows = [
            ("SMTP Server", self.smtp_config["server"], False),
            ("SMTP Port", str(self.smtp_config["port"]), False),
            ("Sender Email", self.smtp_config["sender_email"], False),
            ("App Password", self.smtp_config["sender_password"], True),
            ("Sender Name", self.smtp_config["sender_name"], False),
        ]

        for i, (label, default, is_password) in enumerate(rows):
            ttk.Label(form, text=label).grid(row=i, column=0, sticky="w", pady=6, padx=5)
            entry = ttk.Entry(form, width=44, show="*" if is_password else "")
            entry.insert(0, default)
            entry.grid(row=i, column=1, sticky="w", pady=6, padx=5)
            fields[label] = entry

        pdf_count = sum(1 for r in self.recipients if r.get("pdf_bytes"))
        info_text = (
            f"📎 Each recipient gets their OWN merged PDF attached.\n"
            f"PDFs ready: {pdf_count}/{len(self.recipients)}\n"
            f"🔗 Link columns: "
            f"{', '.join(sorted(self.link_columns)) if self.link_columns else 'none'}\n"
            f"Fallback PDF: {Path(self.attached_pdf_path).name if self.attached_pdf_path else 'None'}"
        )
        ttk.Label(
            form, text=info_text, foreground="#1a6e3c",
            font=("Segoe UI", 9, "italic"), wraplength=500, justify="left"
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=10, padx=5)

        progress_var = tk.DoubleVar()
        progress_label = ttk.Label(win, text="", background="#f0f4f8")
        progress_label.pack(pady=(10, 0))
        progress = ttk.Progressbar(win, variable=progress_var, maximum=100, length=490)
        progress.pack(pady=5)

        button_frame = ttk.Frame(win)
        button_frame.pack(pady=15)

        def do_send():
            server = fields["SMTP Server"].get().strip()
            try:
                port = int(fields["SMTP Port"].get().strip())
            except ValueError:
                messagebox.showerror("Error", "Port must be a number.")
                return

            sender_email = fields["Sender Email"].get().strip()
            sender_password = fields["App Password"].get().strip()
            sender_name = fields["Sender Name"].get().strip() or "Sender"

            if not sender_email or not sender_password:
                messagebox.showerror("Error", "Sender email and App Password are required.")
                return

            self.smtp_config.update({
                "server": server,
                "port": port,
                "sender_email": sender_email,
                "sender_password": sender_password,
                "sender_name": sender_name,
            })

            try:
                self.status.config(text="Connecting to SMTP server...")
                win.update()

                context = ssl.create_default_context()
                if port == 465:
                    smtp = smtplib.SMTP_SSL(server, port, context=context, timeout=30)
                else:
                    smtp = smtplib.SMTP(server, port, timeout=30)
                    smtp.ehlo()
                    smtp.starttls(context=context)
                    smtp.ehlo()

                smtp.login(sender_email, sender_password)
                self.status.config(text="✅ Logged in. Sending...")
                win.update()

                fallback_pdf_bytes = None
                fallback_pdf_name = None
                if self.attached_pdf_path and os.path.exists(self.attached_pdf_path):
                    with open(self.attached_pdf_path, "rb") as f:
                        fallback_pdf_bytes = f.read()
                    fallback_pdf_name = os.path.basename(self.attached_pdf_path)

                total = len(self.recipients)
                sent = 0
                failed = []

                for i, r in enumerate(self.recipients, 1):
                    try:
                        msg = MIMEMultipart("mixed")
                        msg["Subject"] = r["subject"]
                        msg["From"] = formataddr((sender_name, sender_email))
                        msg["To"] = r["email"]
                        if r.get("cc"):
                            msg["Cc"] = r["cc"]

                        html_body = (
                            "<html><body style=\"font-family:Arial,Helvetica,sans-serif;"
                            "font-size:14px;color:#222;line-height:1.6;\">"
                            + r["body"].replace("\n", "<br>\n")
                            + "</body></html>"
                        )
                        plain_body = re.sub(r'<[^>]+>', '', r["body"])

                        alt = MIMEMultipart("alternative")
                        alt.attach(MIMEText(plain_body, "plain", "utf-8"))
                        alt.attach(MIMEText(html_body, "html", "utf-8"))
                        msg.attach(alt)

                        if r.get("pdf_bytes"):
                            part = MIMEApplication(r["pdf_bytes"], _subtype="pdf")
                            part.add_header(
                                "Content-Disposition",
                                "attachment",
                                filename=r["pdf_filename"],
                            )
                            msg.attach(part)
                        elif r.get("docx_bytes"):
                            part = MIMEApplication(
                                r["docx_bytes"],
                                _subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
                            )
                            part.add_header(
                                "Content-Disposition",
                                "attachment",
                                filename=r["docx_filename"],
                            )
                            msg.attach(part)
                        elif fallback_pdf_bytes:
                            part = MIMEApplication(fallback_pdf_bytes, _subtype="pdf")
                            part.add_header(
                                "Content-Disposition",
                                "attachment",
                                filename=fallback_pdf_name,
                            )
                            msg.attach(part)

                        all_recipients = [r["email"]]
                        if r.get("cc"):
                            for c in re.split(r'[;,]', r["cc"]):
                                c = c.strip()
                                if c and "@" in c:
                                    all_recipients.append(c)

                        smtp.sendmail(sender_email, all_recipients, msg.as_string())
                        sent += 1

                    except Exception as e:
                        failed.append((r["email"], str(e)))

                    progress_var.set((i / total) * 100)
                    progress_label.config(text=f"Sending... {i}/{total}")
                    self.status.config(text=f"Sending... {i}/{total}")
                    win.update()

                smtp.quit()

                summary = f"✅ Sent {sent} of {total} email(s)."
                if failed:
                    summary += f"\n\n❌ Failed: {len(failed)}"
                    for email, err in failed[:10]:
                        summary += f"\n  • {email}: {err}"
                    if len(failed) > 10:
                        summary += f"\n  ...and {len(failed) - 10} more."

                messagebox.showinfo("SMTP Send Complete", summary)
                self.status.config(text=f"✅ Sent {sent}/{total} email(s).")
                progress_label.config(text="Done!")
                win.destroy()

            except smtplib.SMTPAuthenticationError:
                messagebox.showerror(
                    "Auth Failed",
                    "SMTP authentication failed.\n\n"
                    "Make sure you're using a Gmail App Password (not your normal password).\n"
                    "Generate one at: https://myaccount.google.com/apppasswords"
                )
                self.status.config(text="❌ SMTP auth failed.")
                progress_label.config(text="Auth failed.")
            except Exception as e:
                messagebox.showerror(
                    "Send Failed",
                    f"Error while sending:\n{str(e)}\n\n{traceback.format_exc()}"
                )
                self.status.config(text="❌ Send failed.")
                progress_label.config(text="Failed.")

        ttk.Button(button_frame, text="🚀 Send All Now", command=do_send,
                   style="Success.TButton").pack(side="left", padx=5)
        ttk.Button(button_frame, text="Cancel", command=win.destroy,
                   style="Primary.TButton").pack(side="left", padx=5)

    # ---------- Debug ----------
    def show_debug_info(self):
        lines = []
        lines.append(f"Word template: {self.word_path or 'not selected'}")
        lines.append(f"PDF engine: {self.pdf_engine or 'NONE'}")
        if self.soffice_path:
            lines.append(f"LibreOffice path: {self.soffice_path}")
        lines.append(f"Placeholders found in Word: {self.placeholders}")
        lines.append(f"Data rows: {len(self.data)}")
        lines.append(f"Data headers: {self.headers}")
        lines.append(f"Link columns: {sorted(self.link_columns) or 'none'}")
        lines.append(f"Email column: {self.email_combo.get() or 'not selected'}")
        lines.append(f"CC column: {self.cc_combo.get() or 'not selected'}")
        lines.append(f"Merged docs: {len(self.merged_docs)}")
        lines.append(f"Recipients: {len(self.recipients)}")

        if self.placeholders and self.headers:
            missing = [p for p in self.placeholders if p not in self.headers]
            unused = [h for h in self.headers if h not in self.placeholders]
            lines.append("")
            lines.append(f"Placeholders missing in data: {missing or 'none'}")
            lines.append(f"Columns not used in Word: {unused or 'none'}")

        if self.merged_docs:
            pdf_ok = sum(1 for d in self.merged_docs if d.get("pdf_bytes"))
            lines.append(f"PDFs generated: {pdf_ok}/{len(self.merged_docs)}")

        if self.recipients:
            lines.append("")
            lines.append("First recipient:")
            r = self.recipients[0]
            lines.append(f"  To: {r['email']}")
            lines.append(f"  CC: {r.get('cc', 'none')}")
            attachment = r.get("pdf_filename") if r.get("pdf_bytes") else r.get("docx_filename")
            lines.append(f"  Attachment: {attachment}")
            lines.append("  Body preview (first 300 chars):")
            lines.append("  " + r["body"][:300].replace("\n", "\n  "))

        messagebox.showinfo("Debug Info", "\n".join(lines))


if __name__ == "__main__":
    root = tk.Tk()
    app = WordMailMergeEmailSender(root)
    root.mainloop()