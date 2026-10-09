import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
import re
import time
import traceback
from pathlib import Path
from datetime import datetime

# yt-dlp is the engine
try:
    import yt_dlp
except ImportError:
    yt_dlp = None


# Browsers yt-dlp can read cookies from
SUPPORTED_BROWSERS = [
    "auto",
    "chrome",
    "edge",
    "firefox",
    "brave",
    "opera",
    "chromium",
    "vivaldi",
    "safari",
]

# Path where a manually exported cookies.txt would live
APP_DIR = os.path.dirname(os.path.abspath(__file__))
COOKIES_TXT_PATH = os.path.join(APP_DIR, "cookies.txt")
DENO_PATH = os.path.join(APP_DIR, "deno.exe")

# Errors that we should retry (transient / YouTube being finicky)
RETRYABLE_PATTERNS = [
    "page needs to be reloaded",
    "unable to extract",
    "failed to extract any player response",
    "signature extraction failed",
    "nsig extraction failed",
    "http error 429",           # rate-limited
    "http error 5",             # 5xx server errors
    "read timed out",
    "connection reset",
    "temporary failure",
    "try again",
]


class YouTubeDownloader:
    def __init__(self, root):
        self.root = root
        self.root.title("🎬 YouTube Downloader")
        self.root.geometry("880x920")
        self.root.configure(bg="#f0f4f8")
        self.root.resizable(True, True)

        # ---------- State ----------
        self.video_info = None
        self.formats = []
        self.save_dir = str(Path.home() / "Downloads")
        self.is_fetching = False
        self.is_downloading = False
        self.is_playlist = False
        self.playlist_count = 0

        # ---------- Main frame ----------
        main = ttk.Frame(root, padding="15")
        main.pack(fill="both", expand=True)

        # Title
        ttk.Label(
            main,
            text="🎬 YouTube Downloader",
            font=("Segoe UI", 16, "bold"),
            foreground="#1a2b3c",
        ).pack(anchor="w")

        ttk.Label(
            main,
            text="Paste a YouTube video or playlist link, pick quality, and download.",
            font=("Segoe UI", 10),
            foreground="#5a6b7c",
        ).pack(anchor="w", pady=(0, 10))

        # ===== URL Row =====
        url_frame = ttk.LabelFrame(main, text="🔗 YouTube URL (video or playlist)", padding=10)
        url_frame.pack(fill="x", pady=5)

        url_row = ttk.Frame(url_frame)
        url_row.pack(fill="x")

        self.url_entry = ttk.Entry(url_row, font=("Segoe UI", 11))
        self.url_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.url_entry.bind("<Return>", lambda e: self.fetch_info())

        self.fetch_btn = ttk.Button(
            url_row,
            text="🔍 Fetch Info",
            command=self.fetch_info,
            style="Primary.TButton",
        )
        self.fetch_btn.pack(side="left", padx=5)

        # Batch mode toggle
        mode_row = ttk.Frame(url_frame)
        mode_row.pack(fill="x", pady=(8, 0))

        self.batch_mode = tk.BooleanVar(value=False)
        self.batch_check = ttk.Checkbutton(
            mode_row,
            text="📚 Batch mode (download all videos in playlist/channel)",
            variable=self.batch_mode,
        )
        self.batch_check.pack(side="left")

        # ===== Cookies / Auth =====
        cookie_frame = ttk.LabelFrame(
            main,
            text="🍪 Authentication (only needed for \"Sign in\" or \"bot check\" errors)",
            padding=10,
        )
        cookie_frame.pack(fill="x", pady=5)

        cookie_row = ttk.Frame(cookie_frame)
        cookie_row.pack(fill="x")

        ttk.Label(cookie_row, text="Browser cookies:").pack(side="left", padx=(0, 8))

        self.browser_var = tk.StringVar(value="auto")
        self.browser_combo = ttk.Combobox(
            cookie_row,
            textvariable=self.browser_var,
            values=SUPPORTED_BROWSERS,
            state="readonly",
            width=15,
        )
        self.browser_combo.pack(side="left", padx=5)

        ttk.Button(
            cookie_row,
            text="🧪 Test Cookies",
            command=self.test_cookies,
            style="Primary.TButton",
        ).pack(side="left", padx=5)

        self.cookie_status = ttk.Label(
            cookie_frame,
            text="",
            font=("Segoe UI", 9),
            foreground="#5a6b7c",
            wraplength=800,
            justify="left",
        )
        self.cookie_status.pack(anchor="w", pady=(6, 0))

        # Auto-detect cookies.txt / deno.exe
        status_parts = []
        if os.path.isfile(COOKIES_TXT_PATH):
            status_parts.append("✅ cookies.txt found")
        else:
            status_parts.append("ℹ️ no cookies.txt (using browser cookies)")

        if os.path.isfile(DENO_PATH):
            status_parts.append("✅ deno.exe found (JS challenges supported)")
        else:
            status_parts.append("ℹ️ no deno.exe (some videos may fail)")

        self.cookie_status.config(
            text="   •   ".join(status_parts),
            foreground="#1a6e3c" if os.path.isfile(COOKIES_TXT_PATH) else "#5a6b7c",
        )

        # ===== Video Info =====
        info_frame = ttk.LabelFrame(main, text="📺 Video Info", padding=10)
        info_frame.pack(fill="x", pady=5)

        self.title_label = ttk.Label(
            info_frame,
            text="No video loaded",
            font=("Segoe UI", 11, "bold"),
            foreground="#1a2b3c",
            wraplength=800,
            justify="left",
        )
        self.title_label.pack(anchor="w", pady=(0, 4))

        self.details_label = ttk.Label(
            info_frame,
            text="",
            font=("Segoe UI", 9),
            foreground="#5a6b7c",
            wraplength=800,
            justify="left",
        )
        self.details_label.pack(anchor="w")

        # ===== Quality =====
        quality_frame = ttk.LabelFrame(main, text="🎚️ Select Quality", padding=10)
        quality_frame.pack(fill="x", pady=5)

        mode_row2 = ttk.Frame(quality_frame)
        mode_row2.pack(fill="x", pady=(0, 8))

        ttk.Label(mode_row2, text="Download type:").pack(side="left", padx=(0, 10))

        self.download_type = tk.StringVar(value="video")
        ttk.Radiobutton(
            mode_row2,
            text="🎥 Video (MP4)",
            variable=self.download_type,
            value="video",
            command=self.refresh_quality_list,
        ).pack(side="left", padx=5)

        ttk.Radiobutton(
            mode_row2,
            text="🎵 Audio (MP3)",
            variable=self.download_type,
            value="audio",
            command=self.refresh_quality_list,
        ).pack(side="left", padx=5)

        quality_row = ttk.Frame(quality_frame)
        quality_row.pack(fill="x", pady=5)

        ttk.Label(quality_row, text="Quality:").pack(side="left", padx=(0, 10))
        self.quality_combo = ttk.Combobox(
            quality_row,
            state="readonly",
            width=55,
            font=("Segoe UI", 10),
        )
        self.quality_combo.pack(side="left", fill="x", expand=True)

        # ===== Save Location =====
        save_frame = ttk.LabelFrame(main, text="💾 Save Location", padding=10)
        save_frame.pack(fill="x", pady=5)

        save_row = ttk.Frame(save_frame)
        save_row.pack(fill="x")

        self.save_label = ttk.Label(
            save_row,
            text=self.save_dir,
            relief="sunken",
            background="white",
            padding=4,
            font=("Segoe UI", 9),
        )
        self.save_label.pack(side="left", fill="x", expand=True, padx=(0, 5))

        ttk.Button(
            save_row,
            text="📂 Change",
            command=self.change_save_dir,
            style="Primary.TButton",
        ).pack(side="left")

        # ===== Progress =====
        progress_frame = ttk.LabelFrame(main, text="📊 Progress", padding=10)
        progress_frame.pack(fill="x", pady=5)

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            length=800,
        )
        self.progress_bar.pack(fill="x")

        self.progress_label = ttk.Label(
            progress_frame,
            text="Ready",
            font=("Segoe UI", 9),
            foreground="#5a6b7c",
        )
        self.progress_label.pack(anchor="w", pady=(4, 0))

        self.batch_progress_label = ttk.Label(
            progress_frame,
            text="",
            font=("Segoe UI", 9),
            foreground="#1a6e3c",
        )
        self.batch_progress_label.pack(anchor="w", pady=(2, 0))

        # ===== Actions =====
        action_frame = ttk.Frame(main)
        action_frame.pack(fill="x", pady=10)

        self.download_btn = ttk.Button(
            action_frame,
            text="⬇️ Download",
            command=self.start_download,
            style="Success.TButton",
        )
        self.download_btn.pack(side="left", padx=5)

        ttk.Button(
            action_frame,
            text="📂 Open Folder",
            command=self.open_save_folder,
            style="Primary.TButton",
        ).pack(side="left", padx=5)

        ttk.Button(
            action_frame,
            text="🧹 Clear",
            command=self.clear_all,
            style="Primary.TButton",
        ).pack(side="left", padx=5)

        # ===== Status bar =====
        self.status = ttk.Label(
            main,
            text="Ready",
            relief="sunken",
            anchor="w",
            background="#e9ecef",
            foreground="#1a2b3c",
            font=("Segoe UI", 9),
            padding=4,
        )
        self.status.pack(fill="x", pady=(5, 0))

        self.setup_styles()

        if yt_dlp is None:
            messagebox.showerror(
                "Missing Library",
                "The 'yt-dlp' library is not installed.\n\n"
                "Install it by running:\n"
                "    pip install yt-dlp\n\n"
                "Then restart this app.",
            )
            self.status.config(text="❌ yt-dlp not installed")

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
        style.configure("Success.TButton", background="#28a745", foreground="white")
        style.map("Success.TButton", background=[("active", "#218838")])
        style.configure("Primary.TButton", background="#2a7de1", foreground="white")
        style.map("Primary.TButton", background=[("active", "#1a5fb4")])

    # ---------- Cookie / auth helpers ----------
    def get_cookie_options(self):
        """
        Return yt-dlp options dict for cookies + JS runtime.
        Priority: cookies.txt > browser cookies
        """
        opts = {}

        # 1. Use cookies.txt if present (most reliable)
        if os.path.isfile(COOKIES_TXT_PATH):
            opts["cookiefile"] = COOKIES_TXT_PATH
        else:
            browser = self.browser_var.get()
            if browser == "auto":
                for b in ("chrome", "edge", "brave", "firefox", "opera", "chromium", "vivaldi"):
                    if self._browser_profile_exists(b):
                        opts["cookiesfrombrowser"] = (b,)
                        break
                else:
                    opts["cookiesfrombrowser"] = ("chrome",)
            else:
                opts["cookiesfrombrowser"] = (browser,)

        # 2. Deno JS runtime if present
        if os.path.isfile(DENO_PATH):
            opts["js_runtimes"] = {"deno": {"path": DENO_PATH}}

        # 3. Make yt-dlp less aggressive about retries at the network level —
        #    we do our own retry loop for cleaner UI feedback.
        opts["retries"] = 2
        opts["fragment_retries"] = 2
        opts["socket_timeout"] = 30
        opts["extractor_retries"] = 1

        return opts

    def _browser_profile_exists(self, browser):
        home = Path.home()
        paths = {
            "chrome": [
                home / "AppData/Local/Google/Chrome/User Data",
                home / ".config/google-chrome",
                home / "Library/Application Support/Google/Chrome",
            ],
            "edge": [
                home / "AppData/Local/Microsoft/Edge/User Data",
                home / "Library/Application Support/Microsoft Edge",
            ],
            "brave": [
                home / "AppData/Local/BraveSoftware/Brave-Browser/User Data",
                home / "Library/Application Support/BraveSoftware/Brave-Browser",
            ],
            "firefox": [
                home / "AppData/Roaming/Mozilla/Firefox/Profiles",
                home / ".mozilla/firefox",
                home / "Library/Application Support/Firefox/Profiles",
            ],
            "opera": [
                home / "AppData/Roaming/Opera Software/Opera Stable",
                home / "Library/Application Support/com.operasoftware.Opera",
            ],
            "chromium": [
                home / "AppData/Local/Chromium/User Data",
                home / ".config/chromium",
            ],
            "vivaldi": [
                home / "AppData/Local/Vivaldi/User Data",
                home / ".config/vivaldi",
            ],
        }
        for p in paths.get(browser, []):
            if p.exists():
                return True
        return False

    def test_cookies(self):
        if yt_dlp is None:
            messagebox.showerror("Missing Library", "yt-dlp is not installed.")
            return

        self.cookie_status.config(
            text="🧪 Testing cookies... please wait (this can take a few seconds).",
            foreground="#5a6b7c",
        )
        self.root.update()

        def run_test():
            try:
                opts = {
                    "quiet": True,
                    "no_warnings": True,
                    "skip_download": True,
                    "noplaylist": True,
                }
                opts.update(self.get_cookie_options())

                with yt_dlp.YoutubeDL(opts) as ydl:
                    ydl.extract_info(
                        "https://www.youtube.com/watch?v=BaW_jenozKc",
                        download=False,
                    )

                self.root.after(0, lambda: self.cookie_status.config(
                    text="✅ Cookies work! You can now fetch and download videos.",
                    foreground="#1a6e3c",
                ))
            except Exception as e:
                err = str(e)[:300]
                self.root.after(0, lambda: self.cookie_status.config(
                    text=f"❌ Cookie test failed: {err}\n\n"
                         f"Try another browser, or export cookies.txt and place it next to this app.",
                    foreground="#d9534f",
                ))

        threading.Thread(target=run_test, daemon=True).start()

    # ---------- Retry helper ----------
    def _is_retryable(self, err_msg: str) -> bool:
        low = err_msg.lower()
        return any(p in low for p in RETRYABLE_PATTERNS)

    def _extract_with_retry(self, url, opts, max_attempts=4, on_status=None):
        """
        Try to extract video info with retries on transient errors.
        Returns the info dict on success, raises on final failure.
        """
        last_error = None
        for attempt in range(1, max_attempts + 1):
            try:
                with yt_dlp.YoutubeDL(opts) as ydl:
                    return ydl.extract_info(url, download=False)
            except Exception as e:
                last_error = e
                msg = str(e)
                if attempt < max_attempts and self._is_retryable(msg):
                    wait = min(2 ** attempt, 10)  # 2s, 4s, 8s
                    if on_status:
                        on_status(f"⚠️ Attempt {attempt} failed. Retrying in {wait}s...")
                    time.sleep(wait)
                    continue
                raise
        if last_error:
            raise last_error

    # ---------- Fetch info ----------
    def fetch_info(self):
        if yt_dlp is None:
            messagebox.showerror("Missing Library", "yt-dlp is not installed.")
            return

        url = self.url_entry.get().strip()
        if not url:
            messagebox.showwarning("Warning", "Please paste a YouTube URL first.")
            return

        if not re.search(r'(youtube\.com|youtu\.be)', url, re.I):
            if not messagebox.askyesno(
                "Not YouTube?",
                "That doesn't look like a YouTube link. Continue anyway?",
            ):
                return

        self.is_fetching = True
        self.fetch_btn.config(state="disabled")
        self.download_btn.config(state="disabled")
        self.title_label.config(text="Fetching info...")
        self.details_label.config(text="")
        self.quality_combo["values"] = []
        self.quality_combo.set("")
        self.status.config(text="🔍 Fetching info...")
        self.progress_var.set(0)
        self.progress_label.config(text="Fetching...")
        self.batch_progress_label.config(text="")

        threading.Thread(target=self._fetch_info_thread, args=(url,), daemon=True).start()

    def _fetch_info_thread(self, url):
        def set_status(msg):
            self.root.after(0, lambda: self.progress_label.config(text=msg))

        try:
            base_opts = {
                "quiet": True,
                "no_warnings": True,
                "skip_download": True,
                "noplaylist": True,
                "extract_flat": False,
            }
            base_opts.update(self.get_cookie_options())

            # Fetch metadata with retry
            info = self._extract_with_retry(url, base_opts, max_attempts=4, on_status=set_status)

            # Detect playlist
            is_playlist = False
            playlist_count = 0
            try:
                probe_opts = {
                    "quiet": True,
                    "no_warnings": True,
                    "skip_download": True,
                    "extract_flat": "in_playlist",
                    "noplaylist": False,
                }
                probe_opts.update(self.get_cookie_options())

                probe = self._extract_with_retry(
                    url, probe_opts, max_attempts=2, on_status=None
                )
                if probe.get("_type") == "playlist":
                    is_playlist = True
                    playlist_count = len(probe.get("entries", []) or [])
            except Exception:
                pass

            self.video_info = info
            self.is_playlist = is_playlist
            self.playlist_count = playlist_count
            self.formats = self.build_format_list(info)

            self.root.after(0, self._show_info_success, info)

        except Exception as e:
            err = str(e)
            self.root.after(0, self._show_info_error, err)

    def _show_info_success(self, info):
        title = info.get("title", "Untitled")
        duration = info.get("duration", 0) or 0
        uploader = info.get("uploader", "Unknown")
        view_count = info.get("view_count", 0) or 0

        mins, secs = divmod(int(duration), 60)
        hours, mins = divmod(mins, 60)
        if hours:
            dur_str = f"{hours}h {mins}m {secs}s"
        else:
            dur_str = f"{mins}m {secs}s"

        self.title_label.config(text=title)
        self.details_label.config(
            text=f"👤 {uploader}   •   ⏱️ {dur_str}   •   👁️ {view_count:,} views"
        )

        if self.is_playlist and self.playlist_count > 0:
            self.batch_progress_label.config(
                text=f"📚 Playlist detected: {self.playlist_count} video(s). "
                     f"Enable batch mode to download all."
            )
            if not self.batch_mode.get():
                self.batch_mode.set(True)
        else:
            self.batch_progress_label.config(text="")
            if self.batch_mode.get() and not self.is_playlist:
                self.batch_mode.set(False)

        self.refresh_quality_list()

        self.is_fetching = False
        self.fetch_btn.config(state="normal")
        self.download_btn.config(state="normal")
        self.status.config(text="✅ Info loaded. Pick a quality and download.")
        self.progress_label.config(text="Ready")

    def _show_info_error(self, err):
        self.is_fetching = False
        self.fetch_btn.config(state="normal")
        self.status.config(text="❌ Failed to fetch info")
        self.progress_label.config(text="Error")

        low = err.lower()

        if "sign in to confirm you" in low or "bot" in low:
            hint = (
                "\n\n💡 TIP: YouTube is asking you to prove you're a real user.\n"
                "   Fix:\n"
                "   1. Log into YouTube in your browser.\n"
                "   2. Pick that browser in the 🍪 Authentication dropdown.\n"
                "   3. Click 🧪 Test Cookies, then try again."
            )
            messagebox.showerror("Bot Check Detected", err + hint)

        elif "could not copy chrome cookie database" in low:
            hint = (
                "\n\n💡 FIX: Fully close Chrome (all windows), then retry.\n"
                "   Or use Firefox, or export a cookies.txt file."
            )
            messagebox.showerror("Chrome Cookie Lock", err + hint)

        elif "page needs to be reloaded" in low:
            hint = (
                "\n\n💡 The retry logic already tried 4 times.\n"
                "   Try:\n"
                "   1. Update yt-dlp:  pip install -U yt-dlp\n"
                "   2. Install Deno next to this app for JS challenges\n"
                "   3. Wait 2-3 minutes and try again"
            )
            messagebox.showerror("YouTube Needs Reload", err + hint)

        else:
            messagebox.showerror("Error", f"Could not fetch video info:\n\n{err}")

    # ---------- Format list ----------
    def build_format_list(self, info):
        video_heights = {}
        audio_formats = []

        for f in info.get("formats", []):
            vcodec = f.get("vcodec") or "none"
            acodec = f.get("acodec") or "none"
            ext = f.get("ext", "")
            height = f.get("height")
            abr = f.get("abr")
            tbr = f.get("tbr")

            if vcodec != "none" and height:
                existing = video_heights.get(height)
                score = (1 if ext == "mp4" else 0, tbr or 0)
                if existing is None or score > existing["score"]:
                    video_heights[height] = {
                        "format_id": f["format_id"],
                        "height": height,
                        "ext": ext,
                        "tbr": tbr,
                        "score": score,
                        "vcodec": vcodec,
                        "acodec": acodec,
                    }

            if vcodec == "none" and acodec != "none":
                audio_formats.append({
                    "format_id": f["format_id"],
                    "ext": ext,
                    "abr": abr or 0,
                    "tbr": tbr or 0,
                })

        video_list = sorted(
            video_heights.values(),
            key=lambda x: x["height"],
            reverse=True,
        )

        options = []
        for v in video_list:
            h = v["height"]
            label = f"🎥 {h}p  ({v['ext'].upper()})"
            if h >= 2160:
                label += "  ⭐ 4K"
            elif h >= 1440:
                label += "  ⭐ 2K"
            elif h >= 1080:
                label += "  ⭐ Full HD"
            elif h >= 720:
                label += "  ⭐ HD"

            options.append({
                "type": "video",
                "format_id": v["format_id"],
                "height": h,
                "ext": v["ext"],
                "label": label,
            })

        audio_formats.sort(key=lambda x: x["abr"] or x["tbr"], reverse=True)
        seen_abrs = set()
        for a in audio_formats:
            abr = int(a["abr"] or a["tbr"] or 0)
            if abr in seen_abrs:
                continue
            seen_abrs.add(abr)
            label = f"🎵 {abr} kbps  ({a['ext'].upper()})"
            options.append({
                "type": "audio",
                "format_id": a["format_id"],
                "abr": abr,
                "ext": a["ext"],
                "label": label,
            })
            if len(seen_abrs) >= 4:
                break

        return options

    def refresh_quality_list(self):
        if not self.formats:
            return

        mode = self.download_type.get()
        filtered = [f for f in self.formats if f["type"] == mode]

        if not filtered:
            self.quality_combo["values"] = ["No formats available"]
            self.quality_combo.set("No formats available")
            return

        labels = [f["label"] for f in filtered]
        self.quality_combo["values"] = labels
        self.quality_combo.current(0)

    def get_selected_format(self):
        selected_label = self.quality_combo.get()
        if not selected_label:
            return None

        mode = self.download_type.get()
        for f in self.formats:
            if f["type"] == mode and f["label"] == selected_label:
                return f
        return None

    def change_save_dir(self):
        folder = filedialog.askdirectory(
            title="Choose where to save downloads",
            initialdir=self.save_dir,
        )
        if folder:
            self.save_dir = folder
            self.save_label.config(text=folder)

    def open_save_folder(self):
        try:
            if not os.path.isdir(self.save_dir):
                os.makedirs(self.save_dir, exist_ok=True)

            if sys.platform == "win32":
                os.startfile(self.save_dir)
            elif sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["open", self.save_dir])
            else:
                import subprocess
                subprocess.Popen(["xdg-open", self.save_dir])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder:\n{e}")

    # ---------- Download ----------
    def start_download(self):
        if yt_dlp is None:
            messagebox.showerror("Missing Library", "yt-dlp is not installed.")
            return

        if not self.video_info:
            messagebox.showwarning("Warning", "Fetch video info first.")
            return

        fmt = self.get_selected_format()
        if not fmt:
            messagebox.showwarning("Warning", "Please select a quality.")
            return

        if self.is_downloading:
            messagebox.showinfo("Info", "A download is already in progress.")
            return

        os.makedirs(self.save_dir, exist_ok=True)

        self.is_downloading = True
        self.download_btn.config(state="disabled")
        self.fetch_btn.config(state="disabled")
        self.progress_var.set(0)
        self.progress_label.config(text="Starting download...")
        self.batch_progress_label.config(text="")
        self.status.config(text="⬇️ Downloading...")

        threading.Thread(
            target=self._download_thread,
            args=(fmt,),
            daemon=True,
        ).start()

    def _download_thread(self, fmt):
        try:
            batch = self.batch_mode.get()
            url = self.video_info.get("webpage_url") or self.url_entry.get().strip()

            if batch:
                playlist_title = self.video_info.get("playlist_title") or self.video_info.get("title") or "playlist"
                playlist_title = re.sub(r'[^\w\-_. ]', '_', playlist_title)[:80]
                out_dir = os.path.join(self.save_dir, playlist_title)
                os.makedirs(out_dir, exist_ok=True)
            else:
                out_dir = self.save_dir

            outtmpl = os.path.join(out_dir, "%(title)s.%(ext)s")

            ydl_opts = {
                "outtmpl": outtmpl,
                "quiet": True,
                "no_warnings": True,
                "noplaylist": not batch,
                "ignoreerrors": True,
                "progress_hooks": [self._progress_hook],
                "postprocessor_hooks": [self._postprocessor_hook],
            }
            ydl_opts.update(self.get_cookie_options())

            if fmt["type"] == "video":
                format_selector = (
                    f"{fmt['format_id']}+bestaudio/"
                    f"{fmt['format_id']}/best"
                )
                ydl_opts["format"] = format_selector
                ydl_opts["merge_output_format"] = "mp4"
            else:
                ydl_opts["format"] = fmt["format_id"]
                ydl_opts["postprocessors"] = [{
                    "key": "FFmpegExtractAudio",
                    "preferredcodec": "mp3",
                    "preferredquality": "192",
                }]

            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                ydl.download([url])

            self.root.after(0, self._download_complete, batch, out_dir)

        except Exception as e:
            err = str(e)
            self.root.after(0, self._download_error, err)

    def _progress_hook(self, d):
        if d.get("status") == "downloading":
            total = d.get("total_bytes") or d.get("total_bytes_estimate")
            downloaded = d.get("downloaded_bytes", 0)
            speed = d.get("speed") or 0
            eta = d.get("eta") or 0

            if total:
                pct = (downloaded / total) * 100
                self.progress_var.set(pct)
            else:
                self.progress_var.set(0)

            speed_str = self._format_bytes(speed) + "/s" if speed else "—"
            eta_str = self._format_time(eta) if eta else "—"
            dl_str = self._format_bytes(downloaded)
            tot_str = self._format_bytes(total) if total else "?"

            info = d.get("info_dict", {})
            title = info.get("title", "")
            playlist_index = info.get("playlist_index")
            playlist_count = info.get("n_entries") or info.get("playlist_count")

            if playlist_index and playlist_count:
                batch_str = f"[{playlist_index}/{playlist_count}]  "
                self.batch_progress_label.config(
                    text=f"📚 Item {playlist_index} of {playlist_count}"
                )
            else:
                batch_str = ""

            self.progress_label.config(
                text=f"{batch_str}{title[:60]}  •  {dl_str} / {tot_str}  •  {speed_str}  •  ETA {eta_str}"
            )

        elif d.get("status") == "finished":
            self.progress_var.set(100)
            self.progress_label.config(text="Download finished. Processing...")

    def _postprocessor_hook(self, d):
        if d.get("status") == "started":
            self.progress_label.config(text="Converting / merging...")
        elif d.get("status") == "finished":
            self.progress_label.config(text="✅ Done with this item!")

    def _download_complete(self, batch, out_dir):
        self.is_downloading = False
        self.download_btn.config(state="normal")
        self.fetch_btn.config(state="normal")
        self.progress_var.set(100)

        if batch:
            self.progress_label.config(text="✅ Batch download complete!")
            self.batch_progress_label.config(
                text=f"📚 All items processed. Saved to: {out_dir}"
            )
            self.status.config(text=f"✅ Batch saved to: {out_dir}")
            msg = f"Batch download complete!\n\nSaved to:\n{out_dir}\n\nOpen folder now?"
        else:
            self.progress_label.config(text="✅ Download complete!")
            self.batch_progress_label.config(text="")
            self.status.config(text=f"✅ Saved to: {out_dir}")
            msg = f"Download complete!\n\nSaved to:\n{out_dir}\n\nOpen folder now?"

        if messagebox.askyesno("Success", msg):
            self.open_save_folder()

    def _download_error(self, err):
        self.is_downloading = False
        self.download_btn.config(state="normal")
        self.fetch_btn.config(state="normal")
        self.progress_label.config(text="❌ Download failed")
        self.status.config(text="❌ Download failed")

        low = err.lower()
        if "sign in to confirm you" in low or "bot" in low:
            hint = (
                "\n\n💡 TIP: YouTube is asking you to prove you're a real user.\n"
                "   Pick your browser in the 🍪 Authentication dropdown, then try again."
            )
            messagebox.showerror("Bot Check Detected", err + hint)
        elif "page needs to be reloaded" in low:
            hint = (
                "\n\n💡 Try:\n"
                "   1. Update yt-dlp:  pip install -U yt-dlp\n"
                "   2. Install Deno next to this app for JS challenges\n"
                "   3. Wait 2-3 minutes and try again"
            )
            messagebox.showerror("YouTube Needs Reload", err + hint)
        else:
            messagebox.showerror("Download Error", err)

    # ---------- Helpers ----------
    def _format_bytes(self, n):
        if not n:
            return "0 B"
        for unit in ("B", "KB", "MB", "GB", "TB"):
            if n < 1024:
                return f"{n:.1f} {unit}"
            n /= 1024
        return f"{n:.1f} PB"

    def _format_time(self, secs):
        if not secs:
            return "0s"
        secs = int(secs)
        m, s = divmod(secs, 60)
        h, m = divmod(m, 60)
        if h:
            return f"{h}h {m}m"
        if m:
            return f"{m}m {s}s"
        return f"{s}s"

    # ---------- Clear ----------
    def clear_all(self):
        self.url_entry.delete(0, tk.END)
        self.video_info = None
        self.formats = []
        self.is_playlist = False
        self.playlist_count = 0
        self.batch_mode.set(False)
        self.title_label.config(text="No video loaded")
        self.details_label.config(text="")
        self.quality_combo["values"] = []
        self.quality_combo.set("")
        self.progress_var.set(0)
        self.progress_label.config(text="Ready")
        self.batch_progress_label.config(text="")
        self.status.config(text="Ready")


if __name__ == "__main__":
    root = tk.Tk()
    app = YouTubeDownloader(root)
    root.mainloop()