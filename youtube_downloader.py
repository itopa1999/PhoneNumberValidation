import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
import re
import traceback
from pathlib import Path
from datetime import datetime

# yt-dlp is the engine
try:
    import yt_dlp
except ImportError:
    yt_dlp = None


class YouTubeDownloader:
    def __init__(self, root):
        self.root = root
        self.root.title("🎬 YouTube Downloader")
        self.root.geometry("850x720")
        self.root.configure(bg="#f0f4f8")
        self.root.resizable(True, True)

        # ---------- State ----------
        self.video_info = None          # raw info dict from yt-dlp
        self.formats = []               # list of {format_id, label, ext, type}
        self.save_dir = str(Path.home() / "Downloads")
        self.is_fetching = False
        self.is_downloading = False

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
            text="Paste a YouTube link, pick your quality, and download.",
            font=("Segoe UI", 10),
            foreground="#5a6b7c",
        ).pack(anchor="w", pady=(0, 10))

        # ===== URL Row =====
        url_frame = ttk.LabelFrame(main, text="🔗 YouTube URL", padding=10)
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

        # ===== Video Info =====
        info_frame = ttk.LabelFrame(main, text="📺 Video Info", padding=10)
        info_frame.pack(fill="x", pady=5)

        self.title_label = ttk.Label(
            info_frame,
            text="No video loaded",
            font=("Segoe UI", 11, "bold"),
            foreground="#1a2b3c",
            wraplength=780,
            justify="left",
        )
        self.title_label.pack(anchor="w", pady=(0, 4))

        self.details_label = ttk.Label(
            info_frame,
            text="",
            font=("Segoe UI", 9),
            foreground="#5a6b7c",
            wraplength=780,
            justify="left",
        )
        self.details_label.pack(anchor="w")

        # ===== Quality Selector =====
        quality_frame = ttk.LabelFrame(main, text="🎚️ Select Quality", padding=10)
        quality_frame.pack(fill="x", pady=5)

        # Mode: Video or Audio
        mode_row = ttk.Frame(quality_frame)
        mode_row.pack(fill="x", pady=(0, 8))

        ttk.Label(mode_row, text="Download type:").pack(side="left", padx=(0, 10))

        self.download_type = tk.StringVar(value="video")
        ttk.Radiobutton(
            mode_row,
            text="🎥 Video (MP4)",
            variable=self.download_type,
            value="video",
            command=self.refresh_quality_list,
        ).pack(side="left", padx=5)

        ttk.Radiobutton(
            mode_row,
            text="🎵 Audio (MP3)",
            variable=self.download_type,
            value="audio",
            command=self.refresh_quality_list,
        ).pack(side="left", padx=5)

        # Quality dropdown
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
            length=780,
        )
        self.progress_bar.pack(fill="x")

        self.progress_label = ttk.Label(
            progress_frame,
            text="Ready",
            font=("Segoe UI", 9),
            foreground="#5a6b7c",
        )
        self.progress_label.pack(anchor="w", pady=(4, 0))

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

        # Warn if yt-dlp is missing
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

    # ---------- Fetch video info ----------
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
        self.status.config(text="🔍 Fetching video info...")
        self.progress_var.set(0)
        self.progress_label.config(text="Fetching...")

        threading.Thread(target=self._fetch_info_thread, args=(url,), daemon=True).start()

    def _fetch_info_thread(self, url):
        try:
            ydl_opts = {
                "quiet": True,
                "no_warnings": True,
                "skip_download": True,
                "noplaylist": True,
            }

            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                info = ydl.extract_info(url, download=False)

            # Playlist?
            if "entries" in info:
                # take first entry
                info = info["entries"][0]

            self.video_info = info

            # Build format list
            self.formats = self.build_format_list(info)

            # Update UI on main thread
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
        messagebox.showerror("Error", f"Could not fetch video info:\n\n{err}")

    # ---------- Build format list ----------
    def build_format_list(self, info):
        """
        Return a clean list of selectable options:
        For video: pick best per height (144p, 240p, 360p, 480p, 720p, 1080p, 1440p, 2160p)
        For audio: pick best audio per bitrate.
        """
        video_heights = {}
        audio_formats = []

        for f in info.get("formats", []):
            vcodec = f.get("vcodec") or "none"
            acodec = f.get("acodec") or "none"
            ext = f.get("ext", "")
            height = f.get("height")
            abr = f.get("abr")
            tbr = f.get("tbr")

            # Video formats
            if vcodec != "none" and height:
                # Prefer mp4 container when available
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

            # Audio-only formats
            if vcodec == "none" and acodec != "none":
                audio_formats.append({
                    "format_id": f["format_id"],
                    "ext": ext,
                    "abr": abr or 0,
                    "tbr": tbr or 0,
                })

        # Sort video by height descending
        video_list = sorted(
            video_heights.values(),
            key=lambda x: x["height"],
            reverse=True,
        )

        # Build video options
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

        # Build audio options (pick top 3 by bitrate)
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

    # ---------- Refresh quality dropdown based on mode ----------
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

    # ---------- Get selected format ----------
    def get_selected_format(self):
        selected_label = self.quality_combo.get()
        if not selected_label:
            return None

        mode = self.download_type.get()
        for f in self.formats:
            if f["type"] == mode and f["label"] == selected_label:
                return f
        return None

    # ---------- Change save dir ----------
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

    # ---------- Start download ----------
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
        self.status.config(text="⬇️ Downloading...")

        threading.Thread(
            target=self._download_thread,
            args=(fmt,),
            daemon=True,
        ).start()

    def _download_thread(self, fmt):
        try:
            # Output template
            outtmpl = os.path.join(
                self.save_dir,
                "%(title)s.%(ext)s",
            )

            ydl_opts = {
                "outtmpl": outtmpl,
                "quiet": True,
                "no_warnings": True,
                "noplaylist": True,
                "progress_hooks": [self._progress_hook],
                "postprocessor_hooks": [self._postprocessor_hook],
            }

            if fmt["type"] == "video":
                # Combine video+audio into mp4 when possible
                format_selector = (
                    f"{fmt['format_id']}+bestaudio/"
                    f"{fmt['format_id']}/best"
                )
                ydl_opts["format"] = format_selector
                ydl_opts["merge_output_format"] = "mp4"

            else:  # audio
                ydl_opts["format"] = fmt["format_id"]
                ydl_opts["postprocessors"] = [{
                    "key": "FFmpegExtractAudio",
                    "preferredcodec": "mp3",
                    "preferredquality": "192",
                }]

            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                ydl.download([self.video_info.get("webpage_url") or self.url_entry.get().strip()])

            self.root.after(0, self._download_complete)

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

            self.progress_label.config(
                text=f"Downloading: {dl_str} / {tot_str}   •   {speed_str}   •   ETA {eta_str}"
            )

        elif d.get("status") == "finished":
            self.progress_var.set(100)
            self.progress_label.config(text="Download finished. Processing...")

    def _postprocessor_hook(self, d):
        if d.get("status") == "started":
            self.progress_label.config(text="Converting / merging...")
        elif d.get("status") == "finished":
            self.progress_label.config(text="✅ Done!")

    def _download_complete(self):
        self.is_downloading = False
        self.download_btn.config(state="normal")
        self.fetch_btn.config(state="normal")
        self.progress_var.set(100)
        self.progress_label.config(text="✅ Download complete!")
        self.status.config(text=f"✅ Saved to: {self.save_dir}")

        if messagebox.askyesno(
            "Success",
            f"Download complete!\n\nSaved to:\n{self.save_dir}\n\nOpen folder now?",
        ):
            self.open_save_folder()

    def _download_error(self, err):
        self.is_downloading = False
        self.download_btn.config(state="normal")
        self.fetch_btn.config(state="normal")
        self.progress_label.config(text="❌ Download failed")
        self.status.config(text="❌ Download failed")
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
        self.title_label.config(text="No video loaded")
        self.details_label.config(text="")
        self.quality_combo["values"] = []
        self.quality_combo.set("")
        self.progress_var.set(0)
        self.progress_label.config(text="Ready")
        self.status.config(text="Ready")


if __name__ == "__main__":
    root = tk.Tk()
    app = YouTubeDownloader(root)
    root.mainloop()