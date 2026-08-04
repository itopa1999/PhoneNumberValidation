import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import re
import webbrowser
import urllib.parse

class WhatsAppLinkGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("📱 WhatsApp Link Generator")
        self.root.geometry("900x700")
        self.root.configure(bg="#f0f4f8")

        # Default message
        self.default_message = (
            "Hello,\n\n"
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

        self.links = []   # store generated links

        # ----- Main container -----
        main = ttk.Frame(root, padding="15")
        main.pack(fill="both", expand=True)

        # Title
        ttk.Label(main, text="📱 WhatsApp Link Generator",
                  font=("Segoe UI", 16, "bold"), foreground="#1a2b3c").pack(anchor="w")
        ttk.Label(main, text="Create wa.me links for multiple phone numbers with a custom message",
                  font=("Segoe UI", 10), foreground="#5a6b7c").pack(anchor="w", pady=(0, 10))

        # ----- Numbers area -----
        num_frame = ttk.LabelFrame(main, text="Phone Numbers (one per line)", padding=10)
        num_frame.pack(fill="both", expand=True, pady=5)

        self.numbers_text = scrolledtext.ScrolledText(num_frame, height=6, font=("Segoe UI", 10))
        self.numbers_text.pack(fill="both", expand=True)

        btn_frame = ttk.Frame(num_frame)
        btn_frame.pack(fill="x", pady=5)
        ttk.Button(btn_frame, text="📂 Upload .txt File", command=self.upload_numbers).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Clear Numbers", command=lambda: self.numbers_text.delete("1.0", tk.END)).pack(side="left", padx=5)

        # ----- Message area -----
        msg_frame = ttk.LabelFrame(main, text="Message (use plain text, will be URL‑encoded)", padding=10)
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

        # ----- Status bar -----
        self.status = ttk.Label(main, text="Ready", relief="sunken", anchor="w",
                                background="#e9ecef", foreground="#1a2b3c", font=("Segoe UI", 9), padding=4)
        self.status.pack(fill="x", pady=(5, 0))

        # Configure styles
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
        style.configure("Danger.TButton", background="#dc3545", foreground="white")
        style.map("Danger.TButton", background=[("active", "#c82333")])
        style.configure("Status.TLabel", background="#e9ecef", foreground="#1a2b3c", font=("Segoe UI", 9), padding=4)
        style.configure("Treeview", background="white", fieldbackground="white", font=("Segoe UI", 9))
        style.map("Treeview", background=[("selected", "#2a7de1")])
        style.configure("Treeview.Heading", background="#d9e2ec", font=("Segoe UI", 10, "bold"))

    # ---------- Number upload ----------
    def upload_numbers(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if not file_path:
            return
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read()
            # Insert into numbers text area
            self.numbers_text.insert(tk.END, content)
            self.status.config(text=f"Loaded numbers from {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not read file:\n{str(e)}")

    # ---------- Generate links ----------
    def generate_links(self):
        # Get numbers
        numbers_raw = self.numbers_text.get("1.0", tk.END).strip()
        if not numbers_raw:
            messagebox.showwarning("Warning", "Please enter phone numbers first.")
            return

        # Split lines and clean each number
        numbers = [line.strip() for line in numbers_raw.splitlines() if line.strip()]
        if not numbers:
            messagebox.showwarning("Warning", "No valid numbers found.")
            return

        # Clean numbers: remove +, spaces, dashes, parentheses, keep digits only
        cleaned_numbers = []
        for num in numbers:
            cleaned = re.sub(r'[^\d]', '', num)
            # Ensure it starts with country code (assume 234 for Nigeria if not present? We'll keep as is)
            # WhatsApp requires international format without '+', just digits.
            if cleaned:
                cleaned_numbers.append(cleaned)

        if not cleaned_numbers:
            messagebox.showwarning("Warning", "No valid phone numbers after cleaning.")
            return

        # Get message
        message = self.msg_text.get("1.0", tk.END).strip()
        if not message:
            messagebox.showwarning("Warning", "Please enter a message.")
            return

        # URL-encode message
        encoded_msg = urllib.parse.quote(message)

        # Generate links
        self.links = [f"https://wa.me/{num}?text={encoded_msg}" for num in cleaned_numbers]

        # Display in listbox
        self.links_listbox.delete(0, tk.END)
        for link in self.links:
            self.links_listbox.insert(tk.END, link)

        self.status.config(text=f"✅ Generated {len(self.links)} link(s).")

    # ---------- Open all in browser ----------
    def open_all(self):
        if not self.links:
            messagebox.showwarning("Warning", "No links generated yet. Click 'Generate Links' first.")
            return
        count = 0
        for link in self.links:
            webbrowser.open_new_tab(link)
            count += 1
        self.status.config(text=f"🌐 Opened {count} link(s) in browser tabs.")

    # ---------- Copy links to clipboard ----------
    def copy_links(self):
        if not self.links:
            messagebox.showwarning("Warning", "No links to copy.")
            return
        text = "\n".join(self.links)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status.config(text=f"📋 Copied {len(self.links)} link(s) to clipboard.")

# ---------- Run ----------
if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppLinkGenerator(root)
    root.mainloop()