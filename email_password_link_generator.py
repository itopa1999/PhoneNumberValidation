import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import re
import webbrowser
import urllib.parse
import pandas as pd
from pathlib import Path
import traceback
import sys
import os
import csv
from datetime import datetime

# SMTP imports
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.utils import formataddr


class EmailLinkGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("📧 Email Link Generator (Excel / CSV)")
        self.root.geometry("1150x950")
        self.root.configure(bg="#f0f4f8")

        # ---------- Default subject ----------
        self.default_subject = (
            "Congratulations! You've Been Selected for the IBS Youth "
            "Entrepreneurial Development Programme (IBS-YEDP)"
        )

        # ---------- Default body ----------
        self.default_body = (
            "Dear {{ firstname }} {{ othername }},\n\n"
            "Congratulations once again, and thank you for enrolling in the "
            "<strong>Ibadan Business School Youth Entrepreneurial Development Programme (IBS-YEDP).</strong>\n\n"
            "We are delighted to welcome you to this transformative three-month "
            "entrepreneurial development and business incubation programme. "
            "Your learning journey officially begins on "
            "<strong>6th October, 2026</strong>, and we look forward to supporting you "
            "as you acquire practical knowledge and skills to build, manage, and grow "
            "a sustainable enterprise.\n\n"
            "<strong>Your Login Credentials</strong>\n\n"
            "Username: {{ email }}\n\n"
            "Password: {{ password }}\n\n"
            "LMS Login Link: https://learn.ibsyedp.school/login/index.php\n\n"
            "Please keep your login credentials confidential and do not share them with anyone.\n\n"
            "<strong>Join the programme's WhatsApp Group</strong>\n\n"
            "Click on this link to join the WhatsApp group where you will get more "
            "information about the programme: "
            "https://chat.whatsapp.com/Es6zmtqJlDtJMAwhNjLXj1?s=cl&p=a&ilr=1\n\n"
            "<strong>Getting Started</strong>\n\n"
            "To access your learning portal:\n\n"
            "1. Click the LMS Login Link above.\n"
            "2. Enter your Username and Password.\n"
            "3. Click <strong>Login</strong>.\n\n"
            "To make the onboarding process easier, we have attached a "
            "<strong>Step-by-Step User Manual (PDF)</strong> to this email. The guide "
            "contains detailed instructions with screenshots to help you log into the "
            "platform, navigate your dashboard, access your courses, join live classes, "
            "submit assignments, and make the most of your learning experience.\n\n"
            "We encourage you to read the manual before your first class.\n\n"
            "<strong>Need Assistance?</strong>\n\n"
            "If you experience any difficulty accessing your account or using the "
            "platform, don't hesitate to get in touch with the IBS-YEDP Support Team "
            "immediately on 09120369651 so we can assist you before the programme commences.\n\n"
            "Once again, welcome to the IBS Youth Entrepreneurial Development Programme.\n\n"
            "We are excited to have you on board and wish you a rewarding learning experience.\n\n"
            "Kind regards,\n\n"
            "<strong>IBS-YEDP Programme Management Team</strong>\n"
            "<strong>Ibadan Business School</strong>\n\n"
            # ================= SIGNATURE START =================
            "<br><br>"
            "<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"vertical-align: -webkit-baseline-middle; font-size: medium; font-family: Arial; min-width: 375px;\">"
            "<tbody>"
            "<tr><td style=\"text-align: center;\"><img role=\"presentation\" width=\"130\" src=\"https://www.dropbox.com/scl/fi/uwhv76os3rb0c1em7233a/YEDP-LOGO.png?rlkey=y4e5qei1njnj2a435bvxxdo0p&st=gxov256b&dl=0&raw=1\" style=\"max-width: 130px; display: inline-block;\"></td></tr>"
            "<tr><td height=\"10\"></td></tr>"
            "<tr style=\"text-align: center;\"><td>"
            "<h2 style=\"margin: 0px; font-size: 18px; font-family: Arial; color: rgb(0, 0, 0); font-weight: 600; line-height: 28px;\">IBS-YEDP Admissions Team</h2>"
            "<p style=\"margin: 0px; color: rgb(0, 0, 0); font-size: 14px; line-height: 22px;\">Youth Entrepreneurship Development Programme (YEDP)</p>"
            "<div style=\"margin: 0px; font-weight: 500; color: rgb(0, 0, 0); font-size: 14px; line-height: 22px;\">School of Enterprise Management | Ibadan Business School (IBS)</div>"
            "</td></tr>"
            "<tr><td>"
            "<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"font-family: Arial; width: 100%;\">"
            "<tbody>"
            "<tr><td height=\"24\"></td></tr>"
            "<tr><td width=\"auto\" style=\"width: 100%; height: 1px; border-bottom: 1px solid rgb(248, 98, 149); display: block;\"></td></tr>"
            "<tr><td height=\"24\"></td></tr>"
            "</tbody></table>"
            "<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"font-family: Arial; width: 100%;\">"
            "<tbody><tr style=\"vertical-align: middle;\">"
            "<td>"
            "<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"font-family: Arial; line-height: 1;\">"
            "<tbody>"
            "<tr style=\"vertical-align: middle; height: 28px;\">"
            "<td width=\"26\" style=\"vertical-align: middle;\"><span style=\"display: inline-block; background-color: rgb(248, 98, 149);\"><img alt=\"mobilePhone\" width=\"18\" src=\"https://cdn2.hubspot.net/hubfs/53/tools/email-signature-generator/icons/phone-icon-dark-2x.png\" style=\"display: block;\"></span></td>"
            "<td style=\"padding: 0px; color: rgb(0, 0, 0);\"><a href=\"tel:+2349120369651\" style=\"text-decoration: none; color: rgb(0, 0, 0); font-size: 14px;\">+2349120369651</a> | <a href=\"tel:+2349120369645\" style=\"text-decoration: none; color: rgb(0, 0, 0); font-size: 14px;\">+2349120369645</a></td>"
            "</tr>"
            "<tr style=\"vertical-align: middle; height: 28px;\">"
            "<td width=\"26\" style=\"vertical-align: middle;\"><span style=\"display: inline-block; background-color: rgb(248, 98, 149);\"><img alt=\"emailAddress\" width=\"18\" src=\"https://cdn2.hubspot.net/hubfs/53/tools/email-signature-generator/icons/email-icon-dark-2x.png\" style=\"display: block;\"></span></td>"
            "<td style=\"padding: 0px; color: rgb(0, 0, 0);\"><a href=\"mailto:ibsyedp@gmail.com\" style=\"text-decoration: none; color: rgb(0, 0, 0); font-size: 14px;\">ibsyedp@gmail.com</a></td>"
            "</tr>"
            "<tr style=\"vertical-align: middle; height: 28px;\">"
            "<td width=\"26\" style=\"vertical-align: middle;\"><span style=\"display: inline-block; background-color: rgb(248, 98, 149);\"><img alt=\"website\" width=\"18\" src=\"https://cdn2.hubspot.net/hubfs/53/tools/email-signature-generator/icons/link-icon-dark-2x.png\" style=\"display: block;\"></span></td>"
            "<td style=\"padding: 0px; color: rgb(0, 0, 0);\"><a href=\"https://www.ibadanbusinessschool.ng\" style=\"text-decoration: none; color: rgb(0, 0, 0); font-size: 14px;\">www.ibadanbusinessschool.ng</a></td>"
            "</tr>"
            "<tr style=\"vertical-align: middle; height: 28px;\">"
            "<td width=\"26\" style=\"vertical-align: middle;\"><span style=\"display: inline-block; background-color: rgb(248, 98, 149);\"><img alt=\"address\" width=\"18\" src=\"https://cdn2.hubspot.net/hubfs/53/tools/email-signature-generator/icons/address-icon-dark-2x.png\" style=\"display: block;\"></span></td>"
            "<td style=\"padding: 0px; color: rgb(0, 0, 0);\"><span style=\"font-size: 14px;\">60 Francis Okediji Street, off Awolowo Avenue, Old Bodija, Ibadan, Oyo State.</span></td>"
            "</tr>"
            "</tbody></table>"
            "</td>"
            "<td width=\"15\"><div style=\"width: 15px;\"></div></td>"
            "<td style=\"text-align: right;\">"
            "<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"font-family: Arial; width: 100%;\">"
            "<tbody>"
            "<tr><td><img role=\"presentation\" width=\"130\" src=\"https://www.dropbox.com/scl/fi/3lbiuudlyceoa7p97uk8g/IBS-Transparent-Logo-crop.png?rlkey=cfj69ymxfflpu6cwwf39keme2&st=486jupbb&dl=0&raw=1\" style=\"max-width: 130px; display: inline-block;\"></td></tr>"
            "<tr><td height=\"10\"></td></tr>"
            "<tr><td>"
            "<div style=\"min-width: 140px;\">"
            "<a href=\"https://www.linkedin.com/company/ibsnigeria/\" style=\"display: inline-block; background-color: rgb(112, 117, 219); border-radius: 50%;\"><img alt=\"linkedin\" width=\"24\" src=\"https://cdn2.hubspot.net/hubfs/53/tools/email-signature-generator/icons/linkedin-icon-dark-2x.png\" style=\"background-color: rgb(112, 117, 219); display: block; border-radius: inherit;\"></a><span style=\"display: inline-block; width: 5px;\"></span>"
            "<a href=\"https://x.com/IBSNigeria\" style=\"display: inline-block; background-color: rgb(112, 117, 219); border-radius: 50%;\"><img alt=\"twitter\" width=\"24\" src=\"https://cdn2.hubspot.net/hubfs/53/tools/email-signature-generator/icons/x-icon-dark-2x.png\" style=\"background-color: rgb(112, 117, 219); display: block; border-radius: inherit;\"></a><span style=\"display: inline-block; width: 5px;\"></span>"
            "<a href=\"https://www.facebook.com/IbadanBusinessSchoolofficial/\" style=\"display: inline-block; background-color: rgb(112, 117, 219); border-radius: 50%;\"><img alt=\"facebook\" width=\"24\" src=\"https://cdn2.hubspot.net/hubfs/53/tools/email-signature-generator/icons/facebook-icon-dark-2x.png\" style=\"background-color: rgb(112, 117, 219); display: block; border-radius: inherit;\"></a><span style=\"display: inline-block; width: 5px;\"></span>"
            "<a href=\"https://www.instagram.com/ibadanbusinessschool/\" style=\"display: inline-block; background-color: rgb(112, 117, 219); border-radius: 50%;\"><img alt=\"instagram\" width=\"24\" src=\"https://cdn2.hubspot.net/hubfs/53/tools/email-signature-generator/icons/instagram-icon-dark-2x.png\" style=\"background-color: rgb(112, 117, 219); display: block; border-radius: inherit;\"></a><span style=\"display: inline-block; width: 5px;\"></span>"
            "<a href=\"https://wa.me/message/JWKJGGXFZ6G4N1\" style=\"display: inline-block; background-color: rgb(112, 117, 219); border-radius: 50%;\"><img alt=\"whatsapp\" width=\"24\" src=\"https://cdn2.hubspot.net/hubfs/53/tools/email-signature-generator/icons/whatsapp-icon-dark-2x.png\" style=\"background-color: rgb(112, 117, 219); display: block; border-radius: inherit;\"></a>"
            "</div>"
            "</td></tr>"
            "</tbody></table>"
            "</td>"
            "</tr></tbody></table>"
            "<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"font-family: Arial; width: 100%;\">"
            "<tbody>"
            "<tr><td height=\"24\"></td></tr>"
            "<tr><td width=\"auto\" style=\"width: 100%; height: 1px; border-bottom: 1px solid rgb(248, 98, 149); display: block;\"></td></tr>"
            "<tr><td height=\"24\"></td></tr>"
            "</tbody></table>"
            "</td></tr>"
            "<tr><td colspan=\"3\" style=\"max-width: 300px; font-size: 12px; padding-top: 1rem; text-align: center;\">"
            "<div class=\"legal-content\">"
            "<p style=\"font-size: inherit; margin: 0px;\">CONFIDENTIALITY NOTICE: This email and any attachments may contain confidential information intended only for the recipient. If received in error, please notify the sender and delete it immediately. Unauthorized use, disclosure, or distribution is prohibited.</p>"
            "<p style=\"font-size: inherit; margin: 0px;\">&nbsp;</p>"
            "<p style=\"font-size: inherit; margin: 0px;\">Registration and participation in the programme are subject to the policies, terms, and admission requirements of Ibadan Business School. IBS reserves the right to accept, decline, or withdraw admission where necessary.</p>"
            "</div>"
            "</td></tr>"
            "</tbody></table>"
            # ================= SIGNATURE END =================
        )
        # ---------- State ----------
        self.data = []
        self.headers = []
        self.email_column = ""
        self.recipients = []
        self.attached_pdf_path = ""   # single PDF for all emails
        self.smtp_config = {
            "server": "smtp.gmail.com",
            "port": 587,
            "sender_email": "ibadanbusinessschool2011@gmail.com",
            "sender_password": "zqoyfnhquoivggut",
            "sender_name": "IBS-YEDP Team",
        }

        # ---------- Main frame ----------
        main = ttk.Frame(root, padding="15")
        main.pack(fill="both", expand=True)

        # Title
        ttk.Label(
            main,
            text="📧 Email Link Generator (Excel / CSV)",
            font=("Segoe UI", 16, "bold"),
            foreground="#1a2b3c",
        ).pack(anchor="w")

        ttk.Label(
            main,
            text="Upload an Excel/CSV file, choose the email column, personalize with {{ ColumnName }} placeholders.",
            font=("Segoe UI", 10),
            foreground="#5a6b7c",
        ).pack(anchor="w", pady=(0, 10))

        # ===== Top row: File upload =====
        top_frame = ttk.Frame(main)
        top_frame.pack(fill="x", pady=5)

        file_frame = ttk.LabelFrame(top_frame, text="Data File (Excel / CSV)", padding=10)
        file_frame.pack(side="left", fill="x", expand=True, padx=(0, 5))

        file_row = ttk.Frame(file_frame)
        file_row.pack(fill="x", pady=2)
        ttk.Label(file_row, text="File:").pack(side="left", padx=5)
        self.file_label = ttk.Label(
            file_row, text="No file selected", relief="sunken",
            background="white", padding=4
        )
        self.file_label.pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(file_row, text="📂 Browse File", command=self.load_file).pack(side="left", padx=5)

        col_row = ttk.Frame(file_frame)
        col_row.pack(fill="x", pady=5)
        ttk.Label(col_row, text="Email column:").pack(side="left", padx=5)
        self.email_combo = ttk.Combobox(col_row, state="readonly", width=30)
        self.email_combo.pack(side="left", padx=5)

        self.preview_label = ttk.Label(file_frame, text="", foreground="#1a6e3c")
        self.preview_label.pack(anchor="w", pady=2)

        # Debug buttons
        debug_frame = ttk.Frame(top_frame)
        debug_frame.pack(side="right", fill="y", padx=5)
        ttk.Button(debug_frame, text="🐞 Debug Info", command=self.show_debug_info,
                   style="Primary.TButton").pack(pady=2)
        ttk.Button(debug_frame, text="📄 Preview First Email",
                   command=self.preview_first_email,
                   style="Primary.TButton").pack(pady=2)

        # ===== PDF attachment row =====
        pdf_frame = ttk.LabelFrame(main, text="📎 Attach PDF (sent with every email)", padding=10)
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

        # ===== Subject =====
        subj_frame = ttk.LabelFrame(main, text="Email Subject", padding=10)
        subj_frame.pack(fill="x", pady=5)

        self.subject_entry = ttk.Entry(subj_frame, font=("Segoe UI", 10))
        self.subject_entry.pack(fill="x")
        self.subject_entry.insert(0, self.default_subject)

        # ===== Message body =====
        msg_frame = ttk.LabelFrame(main, text="Email Body (use {{ ColumnName }} placeholders, HTML <strong>...</strong> allowed)", padding=10)
        msg_frame.pack(fill="both", expand=True, pady=5)

        self.msg_text = scrolledtext.ScrolledText(msg_frame, height=14, font=("Segoe UI", 10), wrap="word")
        self.msg_text.pack(fill="both", expand=True)
        self.msg_text.insert("1.0", self.default_body)

        # ===== Action buttons =====
        action_frame = ttk.Frame(main)
        action_frame.pack(fill="x", pady=10)

        ttk.Button(action_frame, text="⚡ Process (Build Emails)",
                   command=self.process_emails,
                   style="Success.TButton").pack(side="left", padx=5)
        ttk.Button(action_frame, text="🚀 Send via SMTP",
                   command=self.open_smtp_window,
                   style="Success.TButton").pack(side="left", padx=5)
        ttk.Button(action_frame, text="💾 Save as .txt",
                   command=self.save_as_txt,
                   style="Primary.TButton").pack(side="left", padx=5)
        ttk.Button(action_frame, text="💾 Save as .csv",
                   command=self.save_as_csv,
                   style="Primary.TButton").pack(side="left", padx=5)
        ttk.Button(action_frame, text="📋 Copy All",
                   command=self.copy_all,
                   style="Primary.TButton").pack(side="left", padx=5)

        # ===== Results list =====
        links_frame = ttk.LabelFrame(main, text="Processed Emails", padding=10)
        links_frame.pack(fill="both", expand=True, pady=5)

        self.links_listbox = tk.Listbox(links_frame, font=("Segoe UI", 9),
                                        selectmode=tk.EXTENDED)
        self.links_listbox.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(links_frame, orient="vertical",
                                  command=self.links_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.links_listbox.config(yscrollcommand=scrollbar.set)

        self.links_listbox.bind("<Double-Button-1>", self.open_single_email)

        # ===== Status bar =====
        self.status = ttk.Label(
            main, text="Ready", relief="sunken", anchor="w",
            background="#e9ecef", foreground="#1a2b3c",
            font=("Segoe UI", 9), padding=4
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

    # ---------- Load Excel / CSV ----------
    def load_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ]
        )
        if not file_path:
            return

        try:
            lower = file_path.lower()
            if lower.endswith(".csv"):
                try:
                    df = pd.read_csv(file_path, encoding="utf-8")
                except UnicodeDecodeError:
                    df = pd.read_csv(file_path, encoding="latin-1")
            elif lower.endswith(".xls"):
                df = pd.read_excel(file_path, engine="xlrd")
            else:
                df = pd.read_excel(file_path, engine="openpyxl")

            if df.empty:
                messagebox.showerror("Error", "The file is empty.")
                return

            self.data = df.to_dict(orient="records")
            self.headers = list(df.columns)

            self.file_label.config(text=Path(file_path).name)

            email_col = None
            for col in self.headers:
                if re.search(r'email|e-mail|mail', col, re.I):
                    email_col = col
                    break

            self.email_combo["values"] = self.headers
            if email_col:
                self.email_combo.set(email_col)
                self.email_column = email_col
            else:
                self.email_combo.set("")
                self.email_column = ""

            self.preview_label.config(
                text=f"Loaded {len(self.data)} rows, {len(self.headers)} columns."
            )
            self.status.config(
                text=f"Loaded {Path(file_path).name} – {len(self.data)} records"
            )
            self.links_listbox.delete(0, tk.END)
            self.recipients = []

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Failed to read file:\n{str(e)}\n\n{traceback.format_exc()}",
            )
            self.status.config(text="Error loading file")

    # ---------- Attach PDF ----------
    def attach_pdf(self):
        path = filedialog.askopenfilename(
            title="Select a PDF to attach to every email",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not path:
            return
        self.attached_pdf_path = path
        size_kb = os.path.getsize(path) / 1024
        self.pdf_label.config(
            text=f"{Path(path).name}  ({size_kb:.1f} KB)"
        )
        self.status.config(text=f"📎 PDF attached: {Path(path).name}")

    def clear_pdf(self):
        self.attached_pdf_path = ""
        self.pdf_label.config(text="No PDF attached")
        self.status.config(text="PDF attachment cleared.")

    # ---------- Process emails ----------
    def process_emails(self):
        try:
            if not self.data:
                messagebox.showwarning("Warning", "Please load a data file first.")
                return

            self.email_column = self.email_combo.get()
            if not self.email_column:
                messagebox.showwarning("Warning", "Please select the email column.")
                return

            if self.email_column not in self.headers:
                messagebox.showerror("Error", f"Column '{self.email_column}' not found.")
                return

            subject_template = self.subject_entry.get().strip()
            body_template = self.msg_text.get("1.0", tk.END).strip()

            if not subject_template:
                messagebox.showwarning("Warning", "Please enter an email subject.")
                return
            if not body_template:
                messagebox.showwarning("Warning", "Please enter an email body.")
                return

            placeholder_pattern = re.compile(r'\{\{\s*([^}]+?)\s*\}\}')

            self.recipients = []
            skipped = 0
            used_placeholders = set()

            for row in self.data:
                email_raw = str(row.get(self.email_column, "")).strip()
                if not email_raw or "@" not in email_raw:
                    skipped += 1
                    continue

                def replacer(match):
                    key = match.group(1).strip()
                    used_placeholders.add(key)
                    value = row.get(key, "")
                    if value is None or (isinstance(value, float) and pd.isna(value)):
                        value = ""
                    return str(value)

                personalized_subject = placeholder_pattern.sub(replacer, subject_template)
                personalized_body = placeholder_pattern.sub(replacer, body_template)

                first = str(row.get("firstname", "")).strip()
                other = str(row.get("othername", "")).strip()
                display_name = f"{first} {other}".strip() or email_raw

                self.recipients.append({
                    "email": email_raw,
                    "name": display_name,
                    "subject": personalized_subject,
                    "body": personalized_body,
                })

            self.links_listbox.delete(0, tk.END)
            for r in self.recipients:
                self.links_listbox.insert(
                    tk.END,
                    f"{r['name']} <{r['email']}>"
                )

            status_text = f"✅ Processed {len(self.recipients)} email(s)."
            if skipped:
                status_text += f" Skipped {skipped} row(s) with missing/invalid email."
            if used_placeholders:
                status_text += f" Placeholders used: {', '.join(sorted(used_placeholders))}"
            self.status.config(text=status_text)

            if len(self.recipients) == 0:
                self.show_debug_info()

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Error processing emails:\n{str(e)}\n\nTraceback:\n{traceback.format_exc()}",
            )
            self.status.config(text="Error – see popup")

    # ---------- SMTP Window ----------
    def open_smtp_window(self):
        if not self.recipients:
            messagebox.showwarning("Warning", "Process emails first (click 'Process').")
            return

        win = tk.Toplevel(self.root)
        win.title("🚀 Send via SMTP")
        win.geometry("480x520")
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
            entry = ttk.Entry(form, width=38, show="*" if is_password else "")
            entry.insert(0, default)
            entry.grid(row=i, column=1, sticky="w", pady=6, padx=5)
            fields[label] = entry

        # PDF status
        pdf_status = self.attached_pdf_path or "None"
        ttk.Label(form, text="PDF attached:").grid(row=5, column=0, sticky="w", pady=6, padx=5)
        ttk.Label(form, text=Path(pdf_status).name if pdf_status != "None" else "None",
                  foreground="#1a6e3c", font=("Segoe UI", 9, "italic")).grid(
            row=5, column=1, sticky="w", pady=6, padx=5
        )

        # Progress
        progress_var = tk.DoubleVar()
        progress_label = ttk.Label(win, text="", background="#f0f4f8")
        progress_label.pack(pady=(10, 0))
        progress = ttk.Progressbar(win, variable=progress_var, maximum=100, length=400)
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

            # Save to config for next time
            self.smtp_config.update({
                "server": server,
                "port": port,
                "sender_email": sender_email,
                "sender_password": sender_password,
                "sender_name": sender_name,
            })

            # Send
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

                # Read PDF once
                pdf_bytes = None
                pdf_filename = None
                if self.attached_pdf_path and os.path.exists(self.attached_pdf_path):
                    with open(self.attached_pdf_path, "rb") as f:
                        pdf_bytes = f.read()
                    pdf_filename = os.path.basename(self.attached_pdf_path)

                total = len(self.recipients)
                sent = 0
                failed = []

                for i, r in enumerate(self.recipients, 1):
                    try:
                        msg = MIMEMultipart("mixed")
                        msg["Subject"] = r["subject"]
                        msg["From"] = formataddr((sender_name, sender_email))
                        msg["To"] = r["email"]

                        # HTML body (uses <br> for newlines, keeps <strong> bold)
                        html_body = (
                            "<html><body style=\"font-family:Arial,Helvetica,sans-serif;"
                            "font-size:14px;color:#222;line-height:1.6;\">"
                            + r["body"].replace("\n", "<br>\n")
                            + "</body></html>"
                        )

                        # Plain-text fallback (strip HTML tags)
                        plain_body = re.sub(r'<[^>]+>', '', r["body"])

                        alt = MIMEMultipart("alternative")
                        alt.attach(MIMEText(plain_body, "plain", "utf-8"))
                        alt.attach(MIMEText(html_body, "html", "utf-8"))
                        msg.attach(alt)

                        # Attach the same PDF to everyone
                        if pdf_bytes and pdf_filename:
                            part = MIMEApplication(pdf_bytes, _subtype="pdf")
                            part.add_header(
                                "Content-Disposition",
                                "attachment",
                                filename=pdf_filename,
                            )
                            msg.attach(part)

                        smtp.sendmail(sender_email, [r["email"]], msg.as_string())
                        sent += 1

                    except Exception as e:
                        failed.append((r["email"], str(e)))

                    # Update progress
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

    # ---------- Save as .txt (one file per recipient, body only) ----------
    def save_as_txt(self):
        if not self.recipients:
            messagebox.showwarning("Warning", "No processed emails. Click 'Process' first.")
            return

        parent_dir = filedialog.askdirectory(
            title="Select a folder to save all email bodies into",
        )
        if not parent_dir:
            return

        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_dir = Path(parent_dir) / f"emails_{timestamp}"
            output_dir.mkdir(parents=True, exist_ok=True)

            saved_count = 0
            for i, r in enumerate(self.recipients, 1):
                safe_name = re.sub(r'[^\w\-_.@]', '_', r["name"].strip() or "recipient")
                safe_email = re.sub(r'[^\w\-_.@]', '_', r["email"].strip())
                filename = f"{i}_{safe_name}_{safe_email}.txt"
                if len(filename) > 150:
                    filename = filename[:150] + ".txt"

                with open(output_dir / filename, "w", encoding="utf-8") as f:
                    f.write(r["body"])

                saved_count += 1

            self.status.config(text=f"💾 Saved {saved_count} file(s) to {output_dir.name}")

            if messagebox.askyesno(
                "Success",
                f"Saved {saved_count} file(s) to:\n\n{output_dir}\n\nOpen folder now?",
            ):
                try:
                    if sys.platform == "win32":
                        os.startfile(output_dir)
                    elif sys.platform == "darwin":
                        import subprocess
                        subprocess.Popen(["open", str(output_dir)])
                    else:
                        import subprocess
                        subprocess.Popen(["xdg-open", str(output_dir)])
                except Exception as e:
                    messagebox.showwarning("Open Failed", str(e))

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save:\n{str(e)}")

    # ---------- Save as .csv ----------
    def save_as_csv(self):
        if not self.recipients:
            messagebox.showwarning("Warning", "No processed emails. Click 'Process' first.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=f"emails_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        )
        if not file_path:
            return

        try:
            with open(file_path, "w", encoding="utf-8", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["email", "name", "subject", "body"])
                for r in self.recipients:
                    writer.writerow([r["email"], r["name"], r["subject"], r["body"]])

            self.status.config(text=f"💾 Saved to {Path(file_path).name}")
            messagebox.showinfo("Success", f"Saved to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save:\n{str(e)}")

    # ---------- Copy all ----------
    def copy_all(self):
        if not self.recipients:
            messagebox.showwarning("Warning", "No emails to copy.")
            return
        text_parts = []
        for r in self.recipients:
            text_parts.append(f"To: {r['email']}\nSubject: {r['subject']}\n\n{r['body']}")
        full_text = ("\n\n" + "=" * 80 + "\n\n").join(text_parts)
        self.root.clipboard_clear()
        self.root.clipboard_append(full_text)
        self.status.config(text=f"📋 Copied {len(self.recipients)} email(s).")

    # ---------- Open single email (mailto) ----------
    def open_single_email(self, event):
        selection = self.links_listbox.curselection()
        if not selection:
            return
        index = selection[0]
        if index < len(self.recipients):
            r = self.recipients[index]
            try:
                subject_enc = urllib.parse.quote(r["subject"])
                body_enc = urllib.parse.quote(re.sub(r'<[^>]+>', '', r["body"]))
                mailto = f"mailto:{r['email']}?subject={subject_enc}&body={body_enc}"
                webbrowser.open(mailto)
                self.status.config(text=f"Opened draft for {r['email']}")
            except Exception as e:
                self.status.config(text=f"Error: {str(e)[:60]}")

    # ---------- Debug ----------
    def show_debug_info(self):
        if not self.data:
            messagebox.showinfo("Debug", "No data loaded.")
            return

        email_col = self.email_combo.get() or "not selected"
        lines = [
            f"Total rows: {len(self.data)}",
            f"Headers: {self.headers}",
            f"Selected email column: '{email_col}'",
            "",
            "First 5 rows:",
        ]
        for i, row in enumerate(self.data[:5]):
            lines.append(f"Row {i+1}: {row}")

        messagebox.showinfo("Debug Info", "\n".join(lines))

    # ---------- Preview first email ----------
    def preview_first_email(self):
        if not self.recipients:
            messagebox.showinfo("Preview", "No processed emails yet. Click 'Process' first.")
            return
        r = self.recipients[0]
        preview = (
            f"To: {r['email']}\n"
            f"Name: {r['name']}\n\n"
            f"Subject:\n{r['subject']}\n\n"
            f"{'-'*60}\n\n"
            f"{r['body']}"
        )
        messagebox.showinfo("First Email Preview", preview)


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailLinkGenerator(root)
    root.mainloop()