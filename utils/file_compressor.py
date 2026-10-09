import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading
import subprocess
import tempfile
import zipfile
import shutil
from pathlib import Path

from PIL import Image
import pdf2image
import img2pdf

# ---------- Helper Functions ----------
def compress_image(input_path, output_path, quality=85, format=None):
    """Compress or convert an image using Pillow."""
    try:
        with Image.open(input_path) as img:
            if format and format.upper() == 'JPEG':
                if img.mode in ('RGBA', 'P'):
                    img = img.convert('RGB')
            if format:
                save_format = format.upper()
            else:
                save_format = img.format or 'JPEG'
            img.save(output_path, format=save_format, quality=quality, optimize=True)
        return True
    except Exception as e:
        print(f"Image error: {e}")
        return False

def compress_pdf(input_path, output_path, quality=85):
    """Convert PDF to images and recompile with JPEG compression."""
    try:
        images = pdf2image.convert_from_path(input_path, dpi=150)
        with tempfile.TemporaryDirectory() as tmpdir:
            img_paths = []
            for i, img in enumerate(images):
                img_path = os.path.join(tmpdir, f"page_{i+1:03d}.jpg")
                img.save(img_path, 'JPEG', quality=quality, optimize=True)
                img_paths.append(img_path)
            with open(output_path, 'wb') as f:
                f.write(img2pdf.convert(img_paths))
        return True
    except Exception as e:
        print(f"PDF error: {e}")
        return False

def compress_pptx(input_path, output_path, quality=85):
    """Extract PPTX, compress images, repack."""
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            with zipfile.ZipFile(input_path, 'r') as zip_ref:
                zip_ref.extractall(tmpdir)
            media_dir = os.path.join(tmpdir, 'ppt', 'media')
            if os.path.exists(media_dir):
                for fname in os.listdir(media_dir):
                    ext = fname.lower().split('.')[-1]
                    if ext in ('jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff', 'webp'):
                        img_path = os.path.join(media_dir, fname)
                        with Image.open(img_path) as img:
                            img.save(img_path, quality=quality, optimize=True)
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
                for root, _, files in os.walk(tmpdir):
                    for f in files:
                        full_path = os.path.join(root, f)
                        arcname = os.path.relpath(full_path, tmpdir)
                        zip_ref.write(full_path, arcname)
        return True
    except Exception as e:
        print(f"PPTX error: {e}")
        return False

def image_to_pdf(input_paths, output_path, quality=85):
    """Convert one or more images to a single PDF."""
    try:
        images = []
        for path in input_paths:
            with Image.open(path) as img:
                if img.mode == 'RGBA':
                    img = img.convert('RGB')
                with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as tmp:
                    img.save(tmp.name, 'JPEG', quality=quality)
                    images.append(tmp.name)
        with open(output_path, 'wb') as f:
            f.write(img2pdf.convert(images))
        for tmp in images:
            try: os.unlink(tmp)
            except: pass
        return True
    except Exception as e:
        print(f"Image to PDF error: {e}")
        return False

def pdf_to_images(input_path, output_dir, format='JPEG', quality=85):
    """Convert PDF to individual images (one per page)."""
    try:
        images = pdf2image.convert_from_path(input_path, dpi=150)
        out_files = []
        basename = os.path.splitext(os.path.basename(input_path))[0]
        for i, img in enumerate(images):
            out_path = os.path.join(output_dir, f"{basename}_page_{i+1:03d}.{format.lower()}")
            img.save(out_path, format=format, quality=quality, optimize=True)
            out_files.append(out_path)
        return out_files
    except Exception as e:
        print(f"PDF to images error: {e}")
        return []

def image_convert(input_path, output_path, format, quality=85):
    """Convert a single image to another format."""
    try:
        with Image.open(input_path) as img:
            if format.upper() == 'JPEG' and img.mode in ('RGBA', 'P'):
                img = img.convert('RGB')
            img.save(output_path, format=format, quality=quality, optimize=True)
        return True
    except Exception as e:
        print(f"Image convert error: {e}")
        return False

def libreoffice_convert(input_path, output_path, output_format='pdf'):
    """Use LibreOffice to convert DOCX/PPTX to PDF."""
    try:
        subprocess.run([
            'soffice', '--headless', '--convert-to', output_format,
            '--outdir', os.path.dirname(output_path), input_path
        ], check=True, timeout=120, capture_output=True)
        expected = os.path.join(os.path.dirname(output_path),
                                os.path.splitext(os.path.basename(input_path))[0] + '.' + output_format)
        if os.path.exists(expected) and expected != output_path:
            os.rename(expected, output_path)
        return True
    except Exception as e:
        print(f"LibreOffice error: {e}")
        return False

# ---------- Main Application ----------
class UnifiedToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Compressor & Converter")
        self.root.geometry("950x700")
        self.root.minsize(800, 600)
        self.root.configure(bg="#f0f4f8")

        self.setup_styles()

        # Variables shared by both tools
        self.output_dir = tk.StringVar()
        self.quality = tk.IntVar(value=85)
        self.progress = tk.IntVar(value=0)
        self.status_text = tk.StringVar(value="Ready")

        # Tool-specific variables
        # Compression
        self.comp_files = []          # list of input paths
        self.comp_results = []        # list of (input, output, success, size_before, size_after)
        self.comp_save_flags = []     # list of bool

        # Conversion
        self.conv_files = []          # list of input paths
        self.target_format = tk.StringVar(value="PDF")
        self.conv_results = []        # list of (input, output, success, size_before, size_after)
        self.conv_save_flags = []     # list of bool

        # Build UI
        self.build_ui()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        bg = "#f0f4f8"
        fg = "#1a2b3c"
        accent = "#2a7de1"
        success = "#28a745"
        danger = "#dc3545"

        style.configure("TFrame", background=bg)
        style.configure("TLabel", background=bg, foreground=fg, font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6)
        style.map("TButton", background=[("active", accent)])
        style.configure("Success.TButton", background=success, foreground="white")
        style.map("Success.TButton", background=[("active", "#218838")])
        style.configure("Primary.TButton", background=accent, foreground="white")
        style.map("Primary.TButton", background=[("active", "#1a5fb4")])
        style.configure("Danger.TButton", background=danger, foreground="white")
        style.map("Danger.TButton", background=[("active", "#c82333")])
        style.configure("Status.TLabel", background="#e9ecef", foreground=fg, font=("Segoe UI", 9), padding=4)
        style.configure("Treeview", background="white", fieldbackground="white", font=("Segoe UI", 9))
        style.map("Treeview", background=[("selected", accent)])
        style.configure("Treeview.Heading", background="#d9e2ec", font=("Segoe UI", 10, "bold"))
        style.configure("Horizontal.TProgressbar", thickness=20, troughcolor="#e9ecef", background=accent)

    def build_ui(self):
        # Main container
        main = ttk.Frame(self.root, padding="15")
        main.pack(fill="both", expand=True)

        # Header (always visible)
        header = ttk.Frame(main)
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(header, text="🛠️ File Compressor & Converter",
                  font=("Segoe UI", 20, "bold"), foreground="#1a2b3c").pack(side="left")
        ttk.Label(header, text="One tool for compression and format conversion",
                  font=("Segoe UI", 10), foreground="#5a6b7c").pack(side="left", padx=(15, 0))

        ttk.Separator(main, orient="horizontal").pack(fill="x", pady=(0, 10))

        # Navigation bar (tool switcher)
        nav = ttk.Frame(main)
        nav.pack(fill="x", pady=(0, 10))
        self.compress_btn = ttk.Button(nav, text="📦 Compress", command=lambda: self.switch_tool("compress"),
                                       style="Success.TButton")
        self.compress_btn.pack(side="left", padx=(0, 10))
        self.convert_btn = ttk.Button(nav, text="🔄 Convert", command=lambda: self.switch_tool("convert"),
                                      style="Primary.TButton")
        self.convert_btn.pack(side="left")

        # Notebook (tab container) for the two tools
        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill="both", expand=True)

        # Create two tabs
        self.compress_tab = ttk.Frame(self.notebook, padding="10")
        self.convert_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.compress_tab, text="Compress")
        self.notebook.add(self.convert_tab, text="Convert")

        # Build each tab
        self.build_compress_tab()
        self.build_convert_tab()

        # Bottom progress bar (shared)
        progress_frame = ttk.Frame(main)
        progress_frame.pack(fill="x", pady=(10, 0))
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress, maximum=100,
                                            style="Horizontal.TProgressbar")
        self.progress_bar.pack(fill="x")
        self.status_label = ttk.Label(progress_frame, textvariable=self.status_text,
                                      font=("Segoe UI", 9), anchor="w")
        self.status_label.pack(fill="x", pady=(3, 0))

        # Switch to compress by default
        self.switch_tool("compress")

    # ---------- Build Compress Tab ----------
    def build_compress_tab(self):
        parent = self.compress_tab

        # Left panel: file list
        file_frame = ttk.LabelFrame(parent, text="Files to Compress", padding="10")
        file_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        toolbar = ttk.Frame(file_frame)
        toolbar.pack(fill="x", pady=(0, 5))
        ttk.Button(toolbar, text="➕ Add Files", command=self.comp_add_files).pack(side="left", padx=(0, 5))
        ttk.Button(toolbar, text="📁 Add Folder", command=self.comp_add_folder).pack(side="left", padx=(0, 5))
        ttk.Button(toolbar, text="🗑️ Clear", command=self.comp_clear_files, style="Danger.TButton").pack(side="left")

        self.comp_tree = ttk.Treeview(file_frame, columns=("Name", "Type", "Size"), show="headings", height=8)
        self.comp_tree.heading("Name", text="Filename")
        self.comp_tree.heading("Type", text="Type")
        self.comp_tree.heading("Size", text="Size (KB)")
        self.comp_tree.column("Name", width=250)
        self.comp_tree.column("Type", width=80)
        self.comp_tree.column("Size", width=80)
        scroll = ttk.Scrollbar(file_frame, orient="vertical", command=self.comp_tree.yview)
        self.comp_tree.configure(yscrollcommand=scroll.set)
        self.comp_tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        # Right panel: options and preview
        opt_frame = ttk.Frame(parent)
        opt_frame.pack(side="right", fill="both", expand=True, padx=(5, 0))

        # Options
        opt_label = ttk.LabelFrame(opt_frame, text="Compression Options", padding="10")
        opt_label.pack(fill="x", pady=(0, 10))

        quality_frame = ttk.Frame(opt_label)
        quality_frame.pack(fill="x", pady=(5, 0))
        ttk.Label(quality_frame, text="Quality:").pack(side="left", padx=(0, 10))
        self.comp_quality_slider = ttk.Scale(quality_frame, from_=1, to=100, variable=self.quality,
                                             orient="horizontal", length=200)
        self.comp_quality_slider.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.comp_quality_label = ttk.Label(quality_frame, text=f"{self.quality.get()}%", width=6)
        self.comp_quality_label.pack(side="left")
        self.comp_quality_slider.configure(command=lambda v: self.comp_quality_label.config(text=f"{int(float(v))}%"))

        out_frame = ttk.Frame(opt_label)
        out_frame.pack(fill="x", pady=(5, 0))
        ttk.Label(out_frame, text="Output Folder:").pack(side="left", padx=(0, 10))
        out_entry = ttk.Entry(out_frame, textvariable=self.output_dir, width=25)
        out_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(out_frame, text="Browse", command=self.browse_output).pack(side="left")

        # Preview results for compression
        result_label = ttk.LabelFrame(opt_frame, text="Preview & Save", padding="10")
        result_label.pack(fill="both", expand=True)

        self.comp_result_tree = ttk.Treeview(result_label, columns=("Input", "Output", "Before", "After", "Save?"),
                                             show="headings", height=5)
        self.comp_result_tree.heading("Input", text="Input File")
        self.comp_result_tree.heading("Output", text="Output File")
        self.comp_result_tree.heading("Before", text="Before (KB)")
        self.comp_result_tree.heading("After", text="After (KB)")
        self.comp_result_tree.heading("Save?", text="Save?")
        self.comp_result_tree.column("Input", width=120)
        self.comp_result_tree.column("Output", width=120)
        self.comp_result_tree.column("Before", width=70)
        self.comp_result_tree.column("After", width=70)
        self.comp_result_tree.column("Save?", width=60)
        scroll2 = ttk.Scrollbar(result_label, orient="vertical", command=self.comp_result_tree.yview)
        self.comp_result_tree.configure(yscrollcommand=scroll2.set)
        self.comp_result_tree.pack(side="left", fill="both", expand=True)
        scroll2.pack(side="right", fill="y")

        # Action buttons
        btn_frame = ttk.Frame(opt_frame)
        btn_frame.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_frame, text="▶️ Start Compression", command=self.comp_start, style="Success.TButton").pack(side="left", padx=(0, 10))
        ttk.Button(btn_frame, text="💾 Save Selected", command=self.comp_save_selected, style="Primary.TButton").pack(side="left")

        # Bind click to toggle checkmarks
        self.comp_result_tree.bind('<Button-1>', lambda e: self.toggle_save_flag(e, self.comp_result_tree, self.comp_save_flags))

    # ---------- Build Convert Tab ----------
    def build_convert_tab(self):
        parent = self.convert_tab

        # Left panel: file list
        file_frame = ttk.LabelFrame(parent, text="Files to Convert", padding="10")
        file_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        toolbar = ttk.Frame(file_frame)
        toolbar.pack(fill="x", pady=(0, 5))
        ttk.Button(toolbar, text="➕ Add Files", command=self.conv_add_files).pack(side="left", padx=(0, 5))
        ttk.Button(toolbar, text="📁 Add Folder", command=self.conv_add_folder).pack(side="left", padx=(0, 5))
        ttk.Button(toolbar, text="🗑️ Clear", command=self.conv_clear_files, style="Danger.TButton").pack(side="left")

        self.conv_tree = ttk.Treeview(file_frame, columns=("Name", "Type", "Size"), show="headings", height=8)
        self.conv_tree.heading("Name", text="Filename")
        self.conv_tree.heading("Type", text="Type")
        self.conv_tree.heading("Size", text="Size (KB)")
        self.conv_tree.column("Name", width=250)
        self.conv_tree.column("Type", width=80)
        self.conv_tree.column("Size", width=80)
        scroll = ttk.Scrollbar(file_frame, orient="vertical", command=self.conv_tree.yview)
        self.conv_tree.configure(yscrollcommand=scroll.set)
        self.conv_tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        # Right panel: options and preview
        opt_frame = ttk.Frame(parent)
        opt_frame.pack(side="right", fill="both", expand=True, padx=(5, 0))

        # Options
        opt_label = ttk.LabelFrame(opt_frame, text="Conversion Options", padding="10")
        opt_label.pack(fill="x", pady=(0, 10))

        format_frame = ttk.Frame(opt_label)
        format_frame.pack(fill="x", pady=(5, 0))
        ttk.Label(format_frame, text="Output Format:").pack(side="left", padx=(0, 10))
        formats = ["PDF", "JPEG", "PNG", "WEBP", "BMP", "TIFF", "DOCX", "PPTX"]
        self.conv_format_combo = ttk.Combobox(format_frame, textvariable=self.target_format,
                                              values=formats, state="readonly", width=10)
        self.conv_format_combo.pack(side="left")
        ttk.Label(format_frame, text=" (choose target format)").pack(side="left", padx=(10, 0))

        quality_frame = ttk.Frame(opt_label)
        quality_frame.pack(fill="x", pady=(5, 0))
        ttk.Label(quality_frame, text="Quality:").pack(side="left", padx=(0, 10))
        self.conv_quality_slider = ttk.Scale(quality_frame, from_=1, to=100, variable=self.quality,
                                             orient="horizontal", length=200)
        self.conv_quality_slider.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.conv_quality_label = ttk.Label(quality_frame, text=f"{self.quality.get()}%", width=6)
        self.conv_quality_label.pack(side="left")
        self.conv_quality_slider.configure(command=lambda v: self.conv_quality_label.config(text=f"{int(float(v))}%"))

        out_frame = ttk.Frame(opt_label)
        out_frame.pack(fill="x", pady=(5, 0))
        ttk.Label(out_frame, text="Output Folder:").pack(side="left", padx=(0, 10))
        out_entry = ttk.Entry(out_frame, textvariable=self.output_dir, width=25)
        out_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(out_frame, text="Browse", command=self.browse_output).pack(side="left")

        # Preview results for conversion
        result_label = ttk.LabelFrame(opt_frame, text="Preview & Save", padding="10")
        result_label.pack(fill="both", expand=True)

        self.conv_result_tree = ttk.Treeview(result_label, columns=("Input", "Output", "Before", "After", "Save?"),
                                             show="headings", height=5)
        self.conv_result_tree.heading("Input", text="Input File")
        self.conv_result_tree.heading("Output", text="Output File")
        self.conv_result_tree.heading("Before", text="Before (KB)")
        self.conv_result_tree.heading("After", text="After (KB)")
        self.conv_result_tree.heading("Save?", text="Save?")
        self.conv_result_tree.column("Input", width=120)
        self.conv_result_tree.column("Output", width=120)
        self.conv_result_tree.column("Before", width=70)
        self.conv_result_tree.column("After", width=70)
        self.conv_result_tree.column("Save?", width=60)
        scroll2 = ttk.Scrollbar(result_label, orient="vertical", command=self.conv_result_tree.yview)
        self.conv_result_tree.configure(yscrollcommand=scroll2.set)
        self.conv_result_tree.pack(side="left", fill="both", expand=True)
        scroll2.pack(side="right", fill="y")

        # Action buttons
        btn_frame = ttk.Frame(opt_frame)
        btn_frame.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_frame, text="▶️ Start Conversion", command=self.conv_start, style="Success.TButton").pack(side="left", padx=(0, 10))
        ttk.Button(btn_frame, text="💾 Save Selected", command=self.conv_save_selected, style="Primary.TButton").pack(side="left")

        # Bind click to toggle checkmarks
        self.conv_result_tree.bind('<Button-1>', lambda e: self.toggle_save_flag(e, self.conv_result_tree, self.conv_save_flags))

    # ---------- Toggle Save Flag ----------
    def toggle_save_flag(self, event, tree, flag_list):
        region = tree.identify_region(event.x, event.y)
        if region == 'cell':
            col = tree.identify_column(event.x)
            if col == '#5':  # Save? column
                item = tree.identify_row(event.y)
                if item:
                    values = list(tree.item(item, 'values'))
                    if values[4] == '☑':
                        values[4] = '☐'
                    else:
                        values[4] = '☑'
                    tree.item(item, values=values)
                    idx = tree.index(item)
                    flag_list[idx] = (values[4] == '☑')

    # ---------- Common methods ----------
    def browse_output(self):
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_dir.set(folder)

    # ---------- Compression methods ----------
    def comp_add_files(self):
        files = filedialog.askopenfilenames(
            title="Select files to compress",
            filetypes=[("All supported", "*.pdf *.pptx *.jpg *.jpeg *.png *.gif *.bmp *.tiff *.webp"),
                       ("Images", "*.jpg *.jpeg *.png *.gif *.bmp *.tiff *.webp"),
                       ("PDF", "*.pdf"), ("PPTX", "*.pptx")]
        )
        for f in files:
            if f not in self.comp_files:
                self.comp_files.append(f)
                size = os.path.getsize(f) / 1024
                ext = os.path.splitext(f)[1].upper().lstrip('.')
                self.comp_tree.insert("", "end", values=(os.path.basename(f), ext, f"{size:.1f}"))

    def comp_add_folder(self):
        folder = filedialog.askdirectory(title="Select folder")
        if folder:
            count = 0
            for root, _, files in os.walk(folder):
                for f in files:
                    ext = os.path.splitext(f)[1].lower()
                    if ext in ('.pdf', '.pptx', '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp'):
                        full = os.path.join(root, f)
                        if full not in self.comp_files:
                            self.comp_files.append(full)
                            size = os.path.getsize(full) / 1024
                            self.comp_tree.insert("", "end", values=(os.path.basename(full), ext.lstrip('.'), f"{size:.1f}"))
                            count += 1
            if count == 0:
                messagebox.showinfo("No files", "No supported files found.")

    def comp_clear_files(self):
        self.comp_files = []
        for item in self.comp_tree.get_children():
            self.comp_tree.delete(item)
        # clear results
        for item in self.comp_result_tree.get_children():
            self.comp_result_tree.delete(item)
        self.comp_results = []
        self.comp_save_flags = []

    def comp_start(self):
        if not self.comp_files:
            messagebox.showerror("Error", "No files selected.")
            return
        if not self.output_dir.get():
            messagebox.showerror("Error", "Please select an output folder.")
            return

        # Clear previous results
        for item in self.comp_result_tree.get_children():
            self.comp_result_tree.delete(item)
        self.comp_results = []
        self.comp_save_flags = []

        # Disable UI
        self.compress_btn.config(state='disabled')
        self.convert_btn.config(state='disabled')
        self.progress.set(0)
        self.status_text.set("Compressing...")

        threading.Thread(target=self.comp_run, daemon=True).start()

    def comp_run(self):
        total = len(self.comp_files)
        processed = 0
        errors = []
        quality = self.quality.get()
        output_dir = self.output_dir.get()
        os.makedirs(output_dir, exist_ok=True)

        for i, input_path in enumerate(self.comp_files):
            self.status_text.set(f"Compressing {os.path.basename(input_path)} ({i+1}/{total})")
            base = os.path.splitext(os.path.basename(input_path))[0]
            ext = os.path.splitext(input_path)[1].lower()
            out_path = os.path.join(output_dir, f"{base}_compressed{ext}")
            size_before = os.path.getsize(input_path) / 1024
            success = False

            try:
                if ext in ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp'):
                    success = compress_image(input_path, out_path, quality)
                elif ext == '.pdf':
                    success = compress_pdf(input_path, out_path, quality)
                elif ext == '.pptx':
                    success = compress_pptx(input_path, out_path, quality)
                else:
                    errors.append(f"{os.path.basename(input_path)}: Unsupported file type")
                    success = False
            except Exception as e:
                errors.append(f"{os.path.basename(input_path)}: {str(e)}")
                success = False

            size_after = os.path.getsize(out_path) / 1024 if success and os.path.exists(out_path) else 0
            self.comp_results.append((input_path, out_path if success else None, success, size_before, size_after))
            self.comp_save_flags.append(success)

            if success:
                self.comp_result_tree.insert("", "end", values=(os.path.basename(input_path),
                                                                os.path.basename(out_path),
                                                                f"{size_before:.1f}", f"{size_after:.1f}", "☑"))
                processed += 1
            else:
                self.comp_result_tree.insert("", "end", values=(os.path.basename(input_path),
                                                                "❌ Failed",
                                                                f"{size_before:.1f}", "-", "☐"))

            self.progress.set(int((i+1) / total * 100))
            self.status_text.set(f"Processed {i+1}/{total}")

        self.status_text.set(f"Compression complete. {processed} successful, {len(errors)} errors.")
        if errors:
            messagebox.showerror("Errors", "\n".join(errors[:10]) + (f"\n... and {len(errors)-10} more" if len(errors)>10 else ""))

        self.compress_btn.config(state='normal')
        self.convert_btn.config(state='normal')

    def comp_save_selected(self):
        saved = 0
        for i, (inp, out, success, before, after) in enumerate(self.comp_results):
            if self.comp_save_flags[i] and success and out:
                try:
                    shutil.copy2(out, self.output_dir.get())
                    saved += 1
                except:
                    pass
        messagebox.showinfo("Saved", f"{saved} files saved to {self.output_dir.get()}")

    # ---------- Conversion methods ----------
    def conv_add_files(self):
        files = filedialog.askopenfilenames(
            title="Select files to convert",
            filetypes=[("All supported", "*.pdf *.docx *.pptx *.jpg *.jpeg *.png *.gif *.bmp *.tiff *.webp"),
                       ("Documents", "*.pdf *.docx *.pptx"),
                       ("Images", "*.jpg *.jpeg *.png *.gif *.bmp *.tiff *.webp")]
        )
        for f in files:
            if f not in self.conv_files:
                self.conv_files.append(f)
                size = os.path.getsize(f) / 1024
                ext = os.path.splitext(f)[1].upper().lstrip('.')
                self.conv_tree.insert("", "end", values=(os.path.basename(f), ext, f"{size:.1f}"))

    def conv_add_folder(self):
        folder = filedialog.askdirectory(title="Select folder")
        if folder:
            count = 0
            for root, _, files in os.walk(folder):
                for f in files:
                    ext = os.path.splitext(f)[1].lower()
                    if ext in ('.pdf', '.docx', '.pptx', '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp'):
                        full = os.path.join(root, f)
                        if full not in self.conv_files:
                            self.conv_files.append(full)
                            size = os.path.getsize(full) / 1024
                            self.conv_tree.insert("", "end", values=(os.path.basename(full), ext.lstrip('.'), f"{size:.1f}"))
                            count += 1
            if count == 0:
                messagebox.showinfo("No files", "No supported files found.")

    def conv_clear_files(self):
        self.conv_files = []
        for item in self.conv_tree.get_children():
            self.conv_tree.delete(item)
        for item in self.conv_result_tree.get_children():
            self.conv_result_tree.delete(item)
        self.conv_results = []
        self.conv_save_flags = []

    def conv_start(self):
        if not self.conv_files:
            messagebox.showerror("Error", "No files selected.")
            return
        if not self.output_dir.get():
            messagebox.showerror("Error", "Please select an output folder.")
            return

        # Clear previous results
        for item in self.conv_result_tree.get_children():
            self.conv_result_tree.delete(item)
        self.conv_results = []
        self.conv_save_flags = []

        self.compress_btn.config(state='disabled')
        self.convert_btn.config(state='disabled')
        self.progress.set(0)
        self.status_text.set("Converting...")

        threading.Thread(target=self.conv_run, daemon=True).start()

    def conv_run(self):
        total = len(self.conv_files)
        processed = 0
        errors = []
        target = self.target_format.get().lower()
        quality = self.quality.get()
        output_dir = self.output_dir.get()
        os.makedirs(output_dir, exist_ok=True)

        for i, input_path in enumerate(self.conv_files):
            self.status_text.set(f"Converting {os.path.basename(input_path)} ({i+1}/{total})")
            base = os.path.splitext(os.path.basename(input_path))[0]
            input_ext = os.path.splitext(input_path)[1].lower()

            # Determine output extension
            if target == 'pdf':
                out_ext = '.pdf'
            elif target in ('jpg', 'jpeg'):
                out_ext = '.jpg'
            elif target == 'png':
                out_ext = '.png'
            elif target == 'webp':
                out_ext = '.webp'
            elif target == 'bmp':
                out_ext = '.bmp'
            elif target == 'tiff':
                out_ext = '.tiff'
            elif target == 'docx':
                out_ext = '.docx'
            elif target == 'pptx':
                out_ext = '.pptx'

            out_path = os.path.join(output_dir, f"{base}_converted{out_ext}")
            size_before = os.path.getsize(input_path) / 1024
            success = False

            try:
                # Image -> PDF
                if input_ext in ('.jpg','.jpeg','.png','.gif','.bmp','.tiff','.webp') and target == 'pdf':
                    success = image_to_pdf([input_path], out_path, quality)
                # PDF -> Image
                elif input_ext == '.pdf' and target in ('jpg','jpeg','png','webp','bmp','tiff'):
                    out_files = pdf_to_images(input_path, output_dir, target.upper(), quality)
                    if out_files:
                        success = True
                        # Show first page in preview
                        out_path = out_files[0]
                        size_after = sum(os.path.getsize(f) for f in out_files) / 1024
                        self.conv_result_tree.insert("", "end", values=(os.path.basename(input_path),
                                                                        os.path.basename(out_path) + " (multi)",
                                                                        f"{size_before:.1f}", f"{size_after:.1f}", "☑"))
                        self.conv_results.append((input_path, out_files, True, size_before, size_after))
                        self.conv_save_flags.append(True)
                        processed += 1
                        self.progress.set(int((i+1) / total * 100))
                        continue  # skip normal single-file handling
                # Image -> Image
                elif input_ext in ('.jpg','.jpeg','.png','.gif','.bmp','.tiff','.webp') and target in ('jpg','jpeg','png','webp','bmp','tiff'):
                    success = image_convert(input_path, out_path, target.upper(), quality)
                # LibreOffice conversions (DOCX/PPTX -> PDF)
                elif input_ext in ('.docx', '.pptx') and target == 'pdf':
                    try:
                        subprocess.run(['soffice', '--version'], capture_output=True, check=True)
                        success = libreoffice_convert(input_path, out_path, 'pdf')
                    except:
                        errors.append(f"{os.path.basename(input_path)}: LibreOffice not found.")
                        success = False
                else:
                    errors.append(f"{os.path.basename(input_path)}: Conversion from {input_ext} to {target} not supported.")
                    success = False

                if success:
                    size_after = os.path.getsize(out_path) / 1024
                    self.conv_result_tree.insert("", "end", values=(os.path.basename(input_path),
                                                                    os.path.basename(out_path),
                                                                    f"{size_before:.1f}", f"{size_after:.1f}", "☑"))
                    self.conv_results.append((input_path, out_path, True, size_before, size_after))
                    self.conv_save_flags.append(True)
                    processed += 1
                else:
                    self.conv_result_tree.insert("", "end", values=(os.path.basename(input_path),
                                                                    "❌ Failed",
                                                                    f"{size_before:.1f}", "-", "☐"))
                    self.conv_results.append((input_path, None, False, size_before, 0))
                    self.conv_save_flags.append(False)
            except Exception as e:
                errors.append(f"{os.path.basename(input_path)}: {str(e)}")
                self.conv_result_tree.insert("", "end", values=(os.path.basename(input_path),
                                                                "❌ Error",
                                                                f"{size_before:.1f}", "-", "☐"))
                self.conv_results.append((input_path, None, False, size_before, 0))
                self.conv_save_flags.append(False)

            self.progress.set(int((i+1) / total * 100))
            self.status_text.set(f"Processed {i+1}/{total}")

        self.status_text.set(f"Conversion complete. {processed} successful, {len(errors)} errors.")
        if errors:
            messagebox.showerror("Errors", "\n".join(errors[:10]) + (f"\n... and {len(errors)-10} more" if len(errors)>10 else ""))

        self.compress_btn.config(state='normal')
        self.convert_btn.config(state='normal')

    def conv_save_selected(self):
        saved = 0
        for i, (inp, out, success, before, after) in enumerate(self.conv_results):
            if self.conv_save_flags[i] and success:
                if isinstance(out, list):  # multi-page PDF -> images
                    for f in out:
                        try:
                            shutil.copy2(f, self.output_dir.get())
                            saved += 1
                        except:
                            pass
                else:
                    try:
                        shutil.copy2(out, self.output_dir.get())
                        saved += 1
                    except:
                        pass
        messagebox.showinfo("Saved", f"{saved} files saved to {self.output_dir.get()}")

    # ---------- Tool switcher ----------
    def switch_tool(self, tool):
        if tool == "compress":
            self.notebook.select(self.compress_tab)
            self.compress_btn.config(style="Success.TButton")
            self.convert_btn.config(style="Primary.TButton")
        else:
            self.notebook.select(self.convert_tab)
            self.convert_btn.config(style="Success.TButton")
            self.compress_btn.config(style="Primary.TButton")

if __name__ == "__main__":
    root = tk.Tk()
    app = UnifiedToolApp(root)
    root.mainloop()