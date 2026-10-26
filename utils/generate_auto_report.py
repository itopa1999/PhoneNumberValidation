import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import re

# ---------- reportlab imports ----------
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
)
from reportlab.lib.enums import TA_CENTER


# ---------- Custom ttk styles ----------
def setup_styles():
    style = ttk.Style()
    style.theme_use('clam')

    bg = "#f0f4f8"
    fg = "#1a2b3c"
    accent = "#2a7de1"
    success = "#28a745"
    danger = "#dc3545"

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

    style.configure("Nav.TButton",
                    background="#e9ecef", foreground=fg,
                    font=("Segoe UI", 11, "bold"), padding=4)

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

    style.configure("Group.TLabel",
                    background="#dbeafe",
                    foreground="#1a2b3c",
                    font=("Segoe UI", 10, "bold"),
                    padding=6)

    return style


# ---------- Helpers ----------
def sanitize_source(source: str) -> str:
    s = str(source)
    s = s.replace("<br>", " ").replace("<br/>", " ").replace("<BR>", " ")
    s = s.replace("<", " ").replace(">", " ")
    s = re.sub(r'[\\/:*?"|]', "_", s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def build_pdf(source: str, df: pd.DataFrame, week_col: str, out_path: str):
    # Format "Week03" into "Module 03"
    module_name = week_col.replace("Week", "Module ")

    doc = SimpleDocTemplate(
        out_path,
        pagesize=landscape(A4),
        leftMargin=0.5 * inch,
        rightMargin=0.5 * inch,
        topMargin=0.5 * inch,
        bottomMargin=0.5 * inch,
        title=f"Attendance Report - {module_name}",
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "title", parent=styles["Title"],
        fontSize=15, alignment=TA_CENTER,
        textColor=colors.HexColor("#1a2b3c"),
    )
    subtitle_style = ParagraphStyle(
        "sub", parent=styles["Normal"],
        fontSize=11, alignment=TA_CENTER,
        textColor=colors.HexColor("#2a4b6c"),
    )

    elements = []
    # Use module_name instead of week_col for display
    elements.append(Paragraph(f"Attendance Report — {module_name}", title_style))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph(sanitize_source(source), subtitle_style))
    elements.append(Spacer(1, 4))
    elements.append(Paragraph(f"Total candidates: {len(df)}", subtitle_style))
    elements.append(Spacer(1, 12))

    # Use module_name for the table header
    data = [["#", "Name", "Phone", "Type", module_name]]
    for i, (_, row) in enumerate(df.iterrows(), 1):
        data.append([
            str(i),
            str(row.get("Name", "") or ""),
            str(row.get("Phone", "") or ""),
            str(row.get("Type", "") or ""),
            str(row.get(week_col, "") or ""),
        ])

    table = Table(data, repeatRows=1, colWidths=[
        0.4 * inch, 3.6 * inch, 1.5 * inch, 1.2 * inch, 1.3 * inch
    ])
    
    # Base Table Style
    table_style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2a7de1")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), 9),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#b0c0d0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f0f4f8")]),
        ("ALIGN", (0, 1), (0, -1), "CENTER"),
        ("ALIGN", (2, 1), (4, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]

    # Apply colors to the Week/Module status cell dynamically
    for i, (_, row) in enumerate(df.iterrows(), 1):
        status = str(row.get(week_col, "")).strip().lower()
        if status == "present":
            table_style.append(("BACKGROUND", (4, i), (4, i), colors.HexColor("#d4edda"))) # Light Green
            table_style.append(("TEXTCOLOR", (4, i), (4, i), colors.HexColor("#155724"))) # Dark Green text
        elif status == "absent":
            table_style.append(("BACKGROUND", (4, i), (4, i), colors.HexColor("#f8d7da"))) # Light Red
            table_style.append(("TEXTCOLOR", (4, i), (4, i), colors.HexColor("#721c24"))) # Dark Red text

    table.setStyle(TableStyle(table_style))
    elements.append(table)
    doc.build(elements)


# ---------- Main Application ----------
class AttendancePDFApp:
    WEEKS = [f"Week{i:02d}" for i in range(1, 27)]

    def __init__(self, root):
        self.root = root
        self.root.title("Attendance Report Generator")
        self.root.geometry("1100x720")
        self.root.configure(bg="#f0f4f8")

        self.style = setup_styles()

        self.file_path = tk.StringVar()
        self.selected_week = tk.StringVar(value="Week01")

        self.df = None
        self.groups = []
        self.current_index = 0

        main_frame = ttk.Frame(root, padding="15")
        main_frame.pack(fill="both", expand=True)

        # ----- Header -----
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(title_frame,
                  text="📄 Attendance Report Generator",
                  font=("Segoe UI", 18, "bold"),
                  foreground="#1a2b3c").pack(side="left")
        ttk.Label(title_frame,
                  text="Generate per-Source PDF reports for a selected module",
                  font=("Segoe UI", 10),
                  foreground="#5a6b7c").pack(side="left", padx=(15, 0))

        # ----- Description -----
        desc_text = (
            "📌 Steps:\n"
            "   1. Browse and select the Excel file (must contain the 'Attendance_Report' sheet).\n"
            "   2. Choose the week/module (Week01–Week26).\n"
            "   3. Click 'Process' to group candidates by their Source.\n"
            "   4. Use ◀ / ▶ to navigate through groups, then Download the current PDF "
            "or Download All PDFs at once."
        )
        ttk.Label(main_frame, text=desc_text, font=("Segoe UI", 9),
                  foreground="#2a4b6c", background="#f0f4f8",
                  justify="left", wraplength=1050).pack(anchor="w", pady=(0, 10))

        ttk.Separator(main_frame, orient="horizontal").pack(fill="x", pady=(0, 12))

        # ----- File row -----
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill="x", pady=(0, 8))

        ttk.Label(file_frame, text="Excel File:").pack(side="left", padx=(0, 8))
        ttk.Entry(file_frame, textvariable=self.file_path, width=45,
                  font=("Segoe UI", 10)).pack(side="left", padx=(0, 8),
                                              fill="x", expand=True)

        ttk.Button(file_frame, text="📂 Browse",
                   command=self.browse_file).pack(side="left", padx=(0, 6))

        # ----- Week + Action row (with downloads at the top) -----
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill="x", pady=(0, 8))

        ttk.Label(action_frame, text="Week:").pack(side="left", padx=(0, 6))
        self.week_combo = ttk.Combobox(action_frame, textvariable=self.selected_week,
                                       values=self.WEEKS, state="readonly", width=8)
        self.week_combo.pack(side="left", padx=(0, 12))

        ttk.Button(action_frame, text="⚡ Process",
                   command=self.process_file,
                   style="Success.TButton").pack(side="left", padx=(0, 6))

        ttk.Button(action_frame, text="🗑️ Clear",
                   command=self.clear_all,
                   style="Danger.TButton").pack(side="left", padx=(0, 20))

        ttk.Button(action_frame, text="💾 Download Current PDF",
                   command=self.download_current,
                   style="Primary.TButton").pack(side="left", padx=(0, 8))

        ttk.Button(action_frame, text="📦 Download All PDFs",
                   command=self.download_all,
                   style="Primary.TButton").pack(side="left", padx=(0, 8))

        # ----- Group navigation bar -----
        nav_frame = ttk.Frame(main_frame)
        nav_frame.pack(fill="x", pady=(10, 6))

        self.prev_btn = ttk.Button(nav_frame, text="◀ Previous",
                                   command=self.prev_group,
                                   style="Nav.TButton")
        self.prev_btn.pack(side="left")

        self.group_label = ttk.Label(nav_frame, text="No group loaded",
                                     style="Group.TLabel", anchor="center")
        self.group_label.pack(side="left", fill="x", expand=True, padx=10)

        self.next_btn = ttk.Button(nav_frame, text="Next ▶",
                                   command=self.next_group,
                                   style="Nav.TButton")
        self.next_btn.pack(side="right")

        # ----- Status bar -----
        self.status = ttk.Label(main_frame, text="Ready", relief="sunken",
                                anchor="w", style="Status.TLabel")
        self.status.pack(fill="x", pady=(0, 8))

        # ----- Treeview -----
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(
            tree_frame,
            columns=("SN", "Name", "Phone", "Type", "WeekStatus"),
            show="headings",
            height=18,
        )
        self.tree.heading("SN", text="#")
        self.tree.heading("Name", text="Name")
        self.tree.heading("Phone", text="Phone")
        self.tree.heading("Type", text="Type")
        self.tree.heading("WeekStatus", text="Week Status")

        self.tree.column("SN", width=50, anchor="center")
        self.tree.column("Name", width=400, anchor="w")
        self.tree.column("Phone", width=150, anchor="center")
        self.tree.column("Type", width=130, anchor="center")
        self.tree.column("WeekStatus", width=150, anchor="center")

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical",
                                  command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # ----- Bottom bar -----
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill="x", pady=(12, 0))

        ttk.Button(bottom_frame, text="Quit",
                   command=root.quit,
                   style="Danger.TButton").pack(side="right")

        self.count_label = ttk.Label(bottom_frame, text="",
                                     font=("Segoe UI", 9, "italic"))
        self.count_label.pack(side="right", padx=(0, 15))

        self._update_nav_buttons()

    # ---------- Actions ----------
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select Attendance Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if filename:
            self.file_path.set(filename)
            self.status.config(text=f"📁 Loaded: {os.path.basename(filename)}")

    def process_file(self):
        if not self.file_path.get():
            messagebox.showerror("Error", "Please select an Excel file first.")
            return

        week_col = self.selected_week.get()
        if week_col not in self.WEEKS:
            messagebox.showerror("Error", "Please select a valid week.")
            return

        try:
            xls = pd.ExcelFile(self.file_path.get())
            if "Attendance_Report" not in xls.sheet_names:
                messagebox.showerror(
                    "Error",
                    "The Excel file must contain a sheet named 'Attendance_Report'.\n"
                    f"Found sheets: {', '.join(xls.sheet_names)}",
                )
                return

            df = pd.read_excel(self.file_path.get(),
                               sheet_name="Attendance_Report")

            required = ["Name", "Phone", "Type", "Source", week_col]
            missing = [c for c in required if c not in df.columns]
            if missing:
                messagebox.showerror(
                    "Error",
                    f"Missing required columns: {', '.join(missing)}\n"
                    f"(Is the week column '{week_col}' present in the sheet?)",
                )
                return

            df["Name"] = df["Name"].fillna("").astype(str)
            df["Phone"] = df["Phone"].fillna("").astype(str)
            df["Type"] = df["Type"].fillna("").astype(str)
            df["Source"] = df["Source"].fillna("").astype(str)
            df[week_col] = df[week_col].fillna("").astype(str)

            self.df = df

            groups = []
            for source, sub in df.groupby("Source", sort=False):
                # Sort sub-group: "Present" first, "Absent" second
                sub = sub.copy()
                sub['sort_key'] = sub[week_col].apply(
                    lambda x: 0 if str(x).strip().lower() == 'present' else (1 if str(x).strip().lower() == 'absent' else 2)
                )
                sub = sub.sort_values(by='sort_key').drop(columns=['sort_key']).reset_index(drop=True)
                
                groups.append((source, sub))

            groups = [(s, g) for s, g in groups if s.strip()]
            self.groups = groups
            self.current_index = 0
            self._populate_tree()

            self.status.config(
                text=(f"✅ Loaded {len(df)} rows into {len(groups)} Source group(s) "
                      f"for {week_col}")
            )
            self.count_label.config(text=f"{len(df)} rows • {len(groups)} groups")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to process file:\n{e}")
            self.status.config(text="❌ Error")

    def prev_group(self):
        if not self.groups:
            return
        self.current_index = (self.current_index - 1) % len(self.groups)
        self._populate_tree()

    def next_group(self):
        if not self.groups:
            return
        self.current_index = (self.current_index + 1) % len(self.groups)
        self._populate_tree()

    def _populate_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not self.groups:
            self.group_label.config(text="No group loaded")
            self._update_nav_buttons()
            return

        source, sub = self.groups[self.current_index]
        week_col = self.selected_week.get()

        self.group_label.config(
            text=(f"Group {self.current_index + 1} of {len(self.groups)}  —  "
                  f"{sanitize_source(source)}  ({len(sub)} candidate(s))")
        )

        # Update grid header to say "Module 03 Status" instead of Week03
        module_name = week_col.replace("Week", "Module ")
        self.tree.heading("WeekStatus", text=f"{module_name} Status")

        for i, (_, row) in enumerate(sub.iterrows(), 1):
            self.tree.insert(
                "", "end",
                values=(
                    i,
                    row.get("Name", ""),
                    row.get("Phone", ""),
                    row.get("Type", ""),
                    row.get(week_col, ""),
                ),
            )

        self._update_nav_buttons()

    def _update_nav_buttons(self):
        state = "normal" if len(self.groups) > 1 else "disabled"
        self.prev_btn.config(state=state)
        self.next_btn.config(state=state)

    def _current_pdf_name(self, source, week_col):
        safe = sanitize_source(source)
        # Change 'Week03' to 'Module_03' in the filename to avoid spaces
        module_name = week_col.replace("Week", "Module_")
        return f"generated_report_for_{safe}_{module_name}.pdf"

    def download_current(self):
        if not self.groups:
            messagebox.showwarning("Warning", "No data loaded. Process a file first.")
            return

        source, sub = self.groups[self.current_index]
        week_col = self.selected_week.get()
        suggested = self._current_pdf_name(source, week_col)

        out_path = filedialog.asksaveasfilename(
            title="Save Current Group PDF",
            defaultextension=".pdf",
            initialfile=suggested,
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not out_path:
            return

        try:
            build_pdf(source, sub, week_col, out_path)
            self.status.config(text=f"💾 Saved: {os.path.basename(out_path)}")
            messagebox.showinfo("Success", f"PDF saved to:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not create PDF:\n{e}")

    def download_all(self):
        if not self.groups:
            messagebox.showwarning("Warning", "No data loaded. Process a file first.")
            return

        out_dir = filedialog.askdirectory(
            title="Select folder to save all PDF reports"
        )
        if not out_dir:
            return

        week_col = self.selected_week.get()
        success = 0
        errors = []

        for source, sub in self.groups:
            fname = self._current_pdf_name(source, week_col)
            path = os.path.join(out_dir, fname)
            try:
                build_pdf(source, sub, week_col, path)
                success += 1
            except Exception as e:
                errors.append(f"{fname}: {e}")

        msg = f"✅ Generated {success} PDF file(s).\nFolder:\n{out_dir}"
        if errors:
            msg += "\n\n⚠️ Failures:\n" + "\n".join(errors[:5])
        messagebox.showinfo("Done", msg)
        self.status.config(
            text=f"📦 Generated {success}/{len(self.groups)} PDFs in {out_dir}"
        )

    def clear_all(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.df = None
        self.groups = []
        self.current_index = 0
        self.file_path.set("")
        self.group_label.config(text="No group loaded")
        self.status.config(text="Cleared")
        self.count_label.config(text="")
        self._update_nav_buttons()


if __name__ == "__main__":
    root = tk.Tk()
    app = AttendancePDFApp(root)
    root.mainloop()