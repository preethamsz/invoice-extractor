import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk
import threading
import base64
import os
import json
import re
import io
from PIL import Image, ImageTk

# Import logic from our new core folder
from templates.core.pdf_utils import convert_pdf_to_images
from templates.core.ai_extractor import extract_invoice_data
from templates.core.excel_utils import build_excel

# Set the overall theme
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

FIELD_OPTIONS = [
    ("invoice_number", "Invoice Number"),
    ("vendor_name", "Vendor Name"),
    ("invoice_date", "Invoice Date"),
    ("net_amount", "Net Amount"),
    ("tax_amount", "Tax Amount"),
    ("total_amount", "Total Amount"),
]

# ───────────────── Color Palette (InvoiceNet-inspired) ─────────────────
BG_DARK = "#303030"
BG_PANEL = "#2b2b2b"
BG_TOOLBAR = "#383838"
BG_BORDER = "#404040"
BG_CANVAS = "#404040"
BG_CHECKBOX = "#333333"
HIGHLIGHT = "#558de8"
ACCENT = "#e0a050"
TEXT_WHITE = "#ffffff"
TEXT_LIGHT = "#dddddd"
TEXT_DIM = "#aaaaaa"
LOGGER_BG = "#002b36"
LOGGER_FG = "#eee8d5"
BTN_BG = "#484848"
BTN_HOVER = "#558de8"
EXTRACT_BG = "#00897B"
EXTRACT_HOVER = "#26a69a"


class ToolTip:
    """Tooltip on hover for any widget."""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tw = None
        self.widget.bind("<Enter>", self._show)
        self.widget.bind("<Leave>", self._hide)

    def _show(self, event=None):
        x = self.widget.winfo_rootx() + 30
        y = self.widget.winfo_rooty() + 25
        self.tw = tk.Toplevel(self.widget)
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry(f"+{x}+{y}")
        lbl = tk.Label(self.tw, text=self.text, bg="#ffffe0", fg="#333",
                       font=("Segoe UI", 9), relief="solid", borderwidth=1, padx=6, pady=3)
        lbl.pack()

    def _hide(self, event=None):
        if self.tw:
            self.tw.destroy()
            self.tw = None


class IconButton(tk.Button):
    """Dark flat icon button with hover highlight and tooltip."""
    def __init__(self, master, text="", tooltip="", command=None, width=50, height=44, **kw):
        super().__init__(master, text=text, command=command,
                         bg=BG_DARK, fg=TEXT_WHITE, activebackground=HIGHLIGHT,
                         activeforeground=TEXT_WHITE, bd=0, highlightthickness=0,
                         font=("Segoe UI", 14), width=3, relief="flat",
                         cursor="hand2", **kw)
        self.default_bg = BG_DARK
        self.bind("<Enter>", lambda e: self.config(bg=HIGHLIGHT))
        self.bind("<Leave>", lambda e: self.config(bg=self.default_bg))
        if tooltip:
            ToolTip(self, tooltip)


class InvoiceExtractorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Invoice Extractor")
        self.geometry("1360x820")
        self.minsize(1100, 700)

        # State variables
        self.file_path = None
        self.pdf_images = []       # list of base64 strings
        self.pdf_pil_images = []   # list of PIL Image objects for display
        self.current_page = 0
        self.total_pages = 0
        self.result_data = None
        self.zoom_level = 1.0
        self.field_vars = {}
        self.displayed_image = None

        self.build_ui()

    # ═══════════════════════════════════════════════════════
    # UI Construction — 3-Column InvoiceNet-style Layout
    # ═══════════════════════════════════════════════════════
    def build_ui(self):
        self.grid_columnconfigure(0, weight=0)   # left icon toolbar
        self.grid_columnconfigure(1, weight=3)   # center PDF viewer
        self.grid_columnconfigure(2, weight=2)   # right control panel
        self.grid_rowconfigure(0, weight=1)

        self._build_left_toolbar()
        self._build_center_viewer()
        self._build_right_panel()

    # ────────────────── LEFT: Icon Toolbar ──────────────────
    def _build_left_toolbar(self):
        self.left_toolbar = tk.Frame(self, bg=BG_DARK, width=60, bd=0,
                                     highlightbackground=BG_BORDER, highlightthickness=1)
        self.left_toolbar.grid(row=0, column=0, sticky="ns")
        self.left_toolbar.grid_propagate(False)
        self.left_toolbar.configure(width=60)

        self.left_toolbar.columnconfigure(0, weight=1)

        # Top section — main tools
        top_tools = tk.Frame(self.left_toolbar, bg=BG_DARK, bd=0)
        top_tools.pack(side="top", fill="x", pady=(6, 0))

        IconButton(top_tools, text="📂", tooltip="Open File",
                   command=self.browse_file).pack(pady=2, padx=4, fill="x")
        IconButton(top_tools, text="📁", tooltip="Open Directory",
                   command=self.browse_directory).pack(pady=2, padx=4, fill="x")
        IconButton(top_tools, text="💾", tooltip="Save Excel",
                   command=self.save_excel).pack(pady=2, padx=4, fill="x")

        sep1 = tk.Frame(top_tools, bg=BG_BORDER, height=1)
        sep1.pack(fill="x", padx=8, pady=6)

        IconButton(top_tools, text="🧹", tooltip="Clear Page",
                   command=self._clear_page).pack(pady=2, padx=4, fill="x")
        IconButton(top_tools, text="📋", tooltip="Copy JSON",
                   command=self.copy_json_to_clipboard).pack(pady=2, padx=4, fill="x")

        sep2 = tk.Frame(top_tools, bg=BG_BORDER, height=1)
        sep2.pack(fill="x", padx=8, pady=6)

        # File navigation (prev/next file for multi-file)
        file_nav = tk.Frame(top_tools, bg=BG_DARK, bd=0)
        file_nav.pack(pady=2, padx=4, fill="x")
        file_nav.columnconfigure(0, weight=1)
        file_nav.columnconfigure(1, weight=1)

        btn_prev_f = tk.Button(file_nav, text="◀", bg=BG_DARK, fg=TEXT_WHITE,
                               activebackground=HIGHLIGHT, bd=0, font=("Segoe UI", 11),
                               cursor="hand2", command=self.prev_page)
        btn_prev_f.grid(row=0, column=0, sticky="ew")
        btn_next_f = tk.Button(file_nav, text="▶", bg=BG_DARK, fg=TEXT_WHITE,
                               activebackground=HIGHLIGHT, bd=0, font=("Segoe UI", 11),
                               cursor="hand2", command=self.next_page)
        btn_next_f.grid(row=0, column=1, sticky="ew")

        self.file_counter_label = tk.Label(file_nav, text="0 of 0", bg=BG_DARK,
                                           fg=TEXT_DIM, font=("Segoe UI", 8))
        self.file_counter_label.grid(row=1, column=0, columnspan=2, pady=(2, 0))

        # Bottom — help / about
        bottom_tools = tk.Frame(self.left_toolbar, bg=BG_DARK, bd=0)
        bottom_tools.pack(side="bottom", fill="x", pady=(0, 8))

        IconButton(bottom_tools, text="❓", tooltip="Help",
                   command=self._show_help).pack(pady=2, padx=4, fill="x")

    # ────────────────── CENTER: PDF Viewer ──────────────────
    def _build_center_viewer(self):
        center = tk.Frame(self, bg=BG_PANEL, bd=0,
                          highlightbackground=BG_BORDER, highlightthickness=1)
        center.grid(row=0, column=1, sticky="nsew")
        center.grid_rowconfigure(1, weight=1)
        center.grid_columnconfigure(0, weight=1)

        # — Page toolbar —
        page_tools = tk.Frame(center, bg=BG_TOOLBAR, height=36, bd=0,
                              highlightbackground=BG_BORDER, highlightthickness=1)
        page_tools.grid(row=0, column=0, sticky="ew")
        page_tools.grid_columnconfigure(1, weight=1)

        # Navigation cluster (left side)
        nav = tk.Frame(page_tools, bg=BG_TOOLBAR, bd=0)
        nav.grid(row=0, column=0, padx=6, pady=3)

        for txt, cmd in [("⏮", self.first_page), ("◀", self.prev_page)]:
            b = tk.Button(nav, text=txt, bg=BG_TOOLBAR, fg=TEXT_WHITE,
                          activebackground=HIGHLIGHT, bd=0, font=("Segoe UI", 10),
                          cursor="hand2", command=cmd)
            b.pack(side="left", padx=1)
            b.bind("<Enter>", lambda e, w=b: w.config(bg=HIGHLIGHT))
            b.bind("<Leave>", lambda e, w=b: w.config(bg=BG_TOOLBAR))

        self.page_label = tk.Label(nav, text="Page 0 of 0", bg=BG_TOOLBAR,
                                   fg=TEXT_WHITE, font=("Segoe UI", 9))
        self.page_label.pack(side="left", padx=8)

        for txt, cmd in [("▶", self.next_page), ("⏭", self.last_page)]:
            b = tk.Button(nav, text=txt, bg=BG_TOOLBAR, fg=TEXT_WHITE,
                          activebackground=HIGHLIGHT, bd=0, font=("Segoe UI", 10),
                          cursor="hand2", command=cmd)
            b.pack(side="left", padx=1)
            b.bind("<Enter>", lambda e, w=b: w.config(bg=HIGHLIGHT))
            b.bind("<Leave>", lambda e, w=b: w.config(bg=BG_TOOLBAR))

        # Zoom cluster (right side)
        zoom = tk.Frame(page_tools, bg=BG_TOOLBAR, bd=0)
        zoom.grid(row=0, column=1, padx=6, pady=3, sticky="e")

        for txt, cmd in [("🔍+", self.zoom_in), ("🔍−", self.zoom_out)]:
            b = tk.Button(zoom, text=txt, bg=BG_TOOLBAR, fg=TEXT_WHITE,
                          activebackground=HIGHLIGHT, bd=0, font=("Segoe UI", 10),
                          cursor="hand2", command=cmd)
            b.pack(side="left", padx=1)
            b.bind("<Enter>", lambda e, w=b: w.config(bg=HIGHLIGHT))
            b.bind("<Leave>", lambda e, w=b: w.config(bg=BG_TOOLBAR))

        self.zoom_label = tk.Label(zoom, text="100%", bg=BG_TOOLBAR,
                                   fg=TEXT_WHITE, font=("Segoe UI", 9))
        self.zoom_label.pack(side="left", padx=8)

        b_fit = tk.Button(zoom, text="⊞", bg=BG_TOOLBAR, fg=TEXT_WHITE,
                          activebackground=HIGHLIGHT, bd=0, font=("Segoe UI", 11),
                          cursor="hand2", command=self._fit_to_screen)
        b_fit.pack(side="left", padx=1)
        b_fit.bind("<Enter>", lambda e, w=b_fit: w.config(bg=HIGHLIGHT))
        b_fit.bind("<Leave>", lambda e, w=b_fit: w.config(bg=BG_TOOLBAR))
        ToolTip(b_fit, "Fit to Screen")

        # — Canvas —
        canvas_frame = tk.Frame(center, bg=BG_CANVAS, bd=0)
        canvas_frame.grid(row=1, column=0, sticky="nsew")
        canvas_frame.grid_rowconfigure(0, weight=1)
        canvas_frame.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(canvas_frame, bg=BG_CANVAS, highlightthickness=0, cursor="cross")
        self.canvas.grid(row=0, column=0, sticky="nsew")

        self.v_scroll = tk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview,
                                     bg=BG_DARK, highlightbackground=HIGHLIGHT, troughcolor=BG_DARK)
        self.v_scroll.grid(row=0, column=1, sticky="ns")
        self.h_scroll = tk.Scrollbar(canvas_frame, orient="horizontal", command=self.canvas.xview,
                                     bg=BG_DARK, highlightbackground=HIGHLIGHT, troughcolor=BG_DARK)
        self.h_scroll.grid(row=1, column=0, sticky="ew")
        self.canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)

        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Configure>", self._on_canvas_resize)

    # ────────────────── RIGHT: Controls & Results ──────────────────
    def _build_right_panel(self):
        right = tk.Frame(self, bg=BG_DARK, bd=0,
                         highlightbackground=BG_BORDER, highlightthickness=1)
        right.grid(row=0, column=2, sticky="nsew")
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(5, weight=1)  # logger expands

        # ── Logo / Title ──
        logo_frame = tk.Frame(right, bg=BG_DARK, bd=0,
                              highlightbackground=BG_BORDER, highlightthickness=1)
        logo_frame.grid(row=0, column=0, sticky="ew")
        logo_frame.grid_columnconfigure(0, weight=1)

        title_inner = tk.Frame(logo_frame, bg=BG_DARK)
        title_inner.pack(pady=14)
        tk.Label(title_inner, text="⚡", bg=BG_DARK, fg=ACCENT,
                 font=("Segoe UI", 28)).pack(side="left", padx=(0, 8))
        tk.Label(title_inner, text="Invoice Extractor", bg=BG_DARK, fg=TEXT_WHITE,
                 font=("Segoe UI", 22, "bold")).pack(side="left")

        # ── API Key ──
        api_frame = tk.Frame(right, bg=BG_DARK, bd=0, padx=16, pady=8)
        api_frame.grid(row=1, column=0, sticky="ew")
        api_frame.grid_columnconfigure(1, weight=1)

        tk.Label(api_frame, text="API Key:", bg=BG_DARK, fg=TEXT_LIGHT,
                 font=("Segoe UI", 11, "bold")).grid(row=0, column=0, padx=(0, 8), sticky="w")
        self.api_entry = ctk.CTkEntry(api_frame, show="*", placeholder_text="gsk_...",
                                      height=30, border_color="#4a4a6a", fg_color="#1e1e2e",
                                      text_color=TEXT_LIGHT)
        self.api_entry.grid(row=0, column=1, sticky="ew")

        # ── Field Checkboxes ──
        field_outer = tk.Frame(right, bg=BG_DARK, bd=0,
                               highlightbackground=BG_BORDER, highlightthickness=1)
        field_outer.grid(row=2, column=0, sticky="ew", padx=12, pady=(4, 4))

        tk.Label(field_outer, text="Fields:", bg=BG_CHECKBOX, fg=TEXT_WHITE,
                 font=("Segoe UI", 12, "bold"), anchor="w").pack(fill="x", padx=10, pady=(8, 4))

        checkbox_grid = tk.Frame(field_outer, bg=BG_CHECKBOX, bd=0)
        checkbox_grid.pack(fill="x", padx=10, pady=(0, 8))
        checkbox_grid.columnconfigure(0, weight=1)
        checkbox_grid.columnconfigure(1, weight=1)

        for i, (key, label) in enumerate(FIELD_OPTIONS):
            var = tk.BooleanVar(value=True)
            self.field_vars[key] = var
            row_idx = i // 2
            col_idx = i % 2
            cb = tk.Checkbutton(checkbox_grid, text=label, variable=var,
                                bg=BG_CHECKBOX, fg=TEXT_LIGHT, selectcolor=BG_DARK,
                                activebackground=BG_CHECKBOX, activeforeground=TEXT_WHITE,
                                font=("Segoe UI", 11), anchor="w", highlightthickness=0)
            cb.grid(row=row_idx, column=col_idx, sticky="w", padx=(4, 12), pady=2)

        # ── Extract Button (prominent) ──
        extract_frame = tk.Frame(right, bg=BG_DARK, bd=0)
        extract_frame.grid(row=3, column=0, sticky="ew", padx=16, pady=(8, 4))
        extract_frame.grid_columnconfigure(0, weight=1)

        self.extract_btn = tk.Button(extract_frame, text="⚡  Extract", bg=EXTRACT_BG, fg=TEXT_WHITE,
                                     activebackground=EXTRACT_HOVER, activeforeground=TEXT_WHITE,
                                     font=("Segoe UI", 13, "bold"), bd=0, cursor="hand2",
                                     relief="flat", height=2, command=self.start_extraction)
        self.extract_btn.grid(row=0, column=0, sticky="ew")
        self.extract_btn.bind("<Enter>", lambda e: self.extract_btn.config(bg=EXTRACT_HOVER))
        self.extract_btn.bind("<Leave>", lambda e: self.extract_btn.config(bg=EXTRACT_BG))

        # ── Status ──
        self.status_label = tk.Label(right, text="", bg=BG_DARK, fg=TEXT_DIM,
                                     font=("Segoe UI", 10), anchor="w")
        self.status_label.grid(row=4, column=0, sticky="ew", padx=18, pady=(2, 0))

        # ── Logger / Results (Solarized dark) ──
        logger_frame = tk.Frame(right, bg=BG_DARK, bd=0,
                                highlightbackground=BG_BORDER, highlightthickness=1)
        logger_frame.grid(row=5, column=0, sticky="nsew", padx=12, pady=(6, 8))
        logger_frame.grid_rowconfigure(0, weight=1)
        logger_frame.grid_columnconfigure(0, weight=1)

        self.textbox = tk.Text(logger_frame, bg=LOGGER_BG, fg=LOGGER_FG,
                               insertbackground=LOGGER_FG, font=("Consolas", 12),
                               wrap="word", bd=0, relief="flat", padx=10, pady=8)
        self.textbox.grid(row=0, column=0, sticky="nsew")
        log_scroll = tk.Scrollbar(logger_frame, orient="vertical", command=self.textbox.yview,
                                  bg=BG_DARK, troughcolor=LOGGER_BG)
        log_scroll.grid(row=0, column=1, sticky="ns")
        self.textbox.configure(yscrollcommand=log_scroll.set)

        self.textbox.insert("1.0", "Open an invoice PDF or image, then click Extract.\nResults will appear here.")
        self.textbox.configure(state="disabled")

        # ── Bottom Buttons ──
        bottom = tk.Frame(right, bg=BG_DARK, bd=0)
        bottom.grid(row=6, column=0, sticky="ew", padx=12, pady=(0, 12))
        bottom.grid_columnconfigure(0, weight=1)
        bottom.grid_columnconfigure(1, weight=1)

        self.save_btn = tk.Button(bottom, text="💾  Save Information", bg=BTN_BG, fg=TEXT_WHITE,
                                  activebackground=HIGHLIGHT, font=("Segoe UI", 11, "bold"),
                                  bd=0, cursor="hand2", relief="flat", height=2,
                                  state="disabled", command=self.save_excel)
        self.save_btn.grid(row=0, column=0, sticky="ew", padx=(0, 4))

        self.copy_btn = tk.Button(bottom, text="📋  Copy JSON", bg=BTN_BG, fg=TEXT_WHITE,
                                  activebackground=HIGHLIGHT, font=("Segoe UI", 11, "bold"),
                                  bd=0, cursor="hand2", relief="flat", height=2,
                                  state="disabled", command=self.copy_json_to_clipboard)
        self.copy_btn.grid(row=0, column=1, sticky="ew", padx=(4, 0))

    # ════════════════════════════════════════════════════════
    # PDF Viewer Logic
    # ════════════════════════════════════════════════════════
    def _fit_zoom_to_canvas(self):
        if not self.pdf_pil_images:
            return
        img = self.pdf_pil_images[self.current_page]
        canvas_w = self.canvas.winfo_width()
        if canvas_w > 1 and img.width > 0:
            self.zoom_level = (canvas_w - 20) / img.width

    def _fit_to_screen(self):
        if not self.pdf_pil_images:
            return
        if hasattr(self, '_manual_zoom'):
            del self._manual_zoom
        self._fit_zoom_to_canvas()
        self._render_current_page()

    def _on_canvas_resize(self, event=None):
        if self.pdf_pil_images and not hasattr(self, '_manual_zoom'):
            self._fit_zoom_to_canvas()
            self._render_current_page()

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def show_page(self):
        if not self.pdf_pil_images:
            return
        self._fit_zoom_to_canvas()
        if hasattr(self, '_manual_zoom'):
            del self._manual_zoom
        self._render_current_page()

    def _render_current_page(self):
        if not self.pdf_pil_images:
            return
        img = self.pdf_pil_images[self.current_page]
        w = int(img.width * self.zoom_level)
        h = int(img.height * self.zoom_level)
        resized = img.resize((w, h), Image.LANCZOS)
        self.displayed_image = ImageTk.PhotoImage(resized)

        canvas_w = self.canvas.winfo_width()
        x_offset = max((canvas_w - w) // 2, 0)

        self.canvas.delete("all")
        self.canvas.create_rectangle(x_offset, 10, x_offset + w, 10 + h, fill="white", outline="#555")
        self.canvas.create_image(x_offset, 10, anchor="nw", image=self.displayed_image)
        self.canvas.configure(scrollregion=(0, 0, max(canvas_w, w + 20), h + 20))

        self.page_label.configure(text=f"Page {self.current_page + 1} of {self.total_pages}")
        self.zoom_label.configure(text=f"{int(self.zoom_level * 100)}%")
        self.file_counter_label.configure(text=f"{self.current_page + 1} of {self.total_pages}")

    # ────────────────── Page Navigation ──────────────────
    def first_page(self):
        if self.total_pages:
            self.current_page = 0
            self.show_page()

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.show_page()

    def next_page(self):
        if self.current_page < self.total_pages - 1:
            self.current_page += 1
            self.show_page()

    def last_page(self):
        if self.total_pages:
            self.current_page = self.total_pages - 1
            self.show_page()

    def zoom_in(self):
        self._manual_zoom = True
        self.zoom_level = min(self.zoom_level + 0.15, 5.0)
        self._render_current_page()

    def zoom_out(self):
        self._manual_zoom = True
        self.zoom_level = max(self.zoom_level - 0.15, 0.2)
        self._render_current_page()

    # ────────────────── File Handling ──────────────────
    def browse_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Supported Files", "*.pdf *.jpg *.jpeg *.png *.webp"),
                       ("PDF Files", "*.pdf"),
                       ("Image Files", "*.jpg *.jpeg *.png *.webp")])
        if not path:
            return
        self._load_file(path)

    def browse_directory(self):
        dir_path = filedialog.askdirectory(title="Select Directory Containing Invoices")
        if not dir_path:
            return
        supported = {'.pdf', '.jpg', '.jpeg', '.png', '.webp'}
        files = [os.path.join(dir_path, f) for f in sorted(os.listdir(dir_path))
                 if os.path.splitext(f)[1].lower() in supported]
        if not files:
            messagebox.showinfo("No Files", "No supported files found in this directory.")
            return
        self._load_file(files[0])
        self._log(f"Found {len(files)} file(s) in directory.")

    def _load_file(self, path):
        self.file_path = path
        ext = os.path.splitext(path)[1].lower()

        if ext == ".pdf":
            self.status_label.configure(text="Converting PDF...", fg="#ff8c69")
            threading.Thread(target=self.process_pdf, args=(path,), daemon=True).start()
        else:
            self.pdf_images = []
            img = Image.open(path)
            self.pdf_pil_images = [img]
            self.total_pages = 1
            self.current_page = 0
            with open(path, "rb") as f:
                self.pdf_images = [base64.b64encode(f.read()).decode()]
            self.zoom_level = 1.0
            self.show_page()
            self.status_label.configure(text=f"Loaded: {os.path.basename(path)}", fg="#4ecb8d")
            self._log(f"Loaded image '{os.path.basename(path)}'")

    def process_pdf(self, path):
        try:
            self.pdf_images = convert_pdf_to_images(path)
            pil_imgs = []
            for b64 in self.pdf_images:
                raw = base64.b64decode(b64)
                pil_imgs.append(Image.open(io.BytesIO(raw)))
            self.pdf_pil_images = pil_imgs
            self.total_pages = len(pil_imgs)
            self.current_page = 0
            self.zoom_level = 1.0
            self.after(0, self.show_page)
            self.after(0, lambda: self.status_label.configure(
                text=f"Ready: {self.total_pages} page(s)", fg="#4ecb8d"))
            self.after(0, lambda: self._log(f"Loaded PDF '{os.path.basename(path)}' — {self.total_pages} page(s)"))
        except Exception as e:
            self.after(0, lambda: self.status_label.configure(text="Error converting PDF", fg="#e63946"))
            self.after(0, lambda: messagebox.showerror("PDF Error", str(e)))

    # ────────────────── Extraction ──────────────────
    def start_extraction(self):
        api_key = self.api_entry.get().strip()
        if not api_key:
            messagebox.showwarning("Missing API Key", "Please enter your Groq API Key.")
            return

        if not self.pdf_images:
            messagebox.showwarning("No Data", "Please open a PDF or image file first.")
            return

        self.extract_btn.configure(state="disabled", text="⏳  Processing...")
        self.save_btn.configure(state="disabled")
        self.copy_btn.configure(state="disabled")
        self.status_label.configure(text="Analyzing via Groq AI...", fg=TEXT_WHITE)
        self._log("Extracting information from current page...")

        threading.Thread(target=self.run_extraction, args=(api_key,), daemon=True).start()

    def run_extraction(self, api_key):
        try:
            image_b64 = self.pdf_images[self.current_page]
            ext = os.path.splitext(self.file_path)[1].lower() if self.file_path else ".png"
            mime_map = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".webp": "image/webp"}
            mime = mime_map.get(ext, "image/png")

            self.result_data = extract_invoice_data(api_key, image_b64, mime)
            self.after(0, self.display_results)
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Extraction Error", str(e)))
        finally:
            self.after(0, self.reset_ui_state)

    def display_results(self):
        self.save_btn.configure(state="normal")
        self.copy_btn.configure(state="normal")

        # Filter to only selected fields
        selected = {k for k, v in self.field_vars.items() if v.get()}
        if selected and self.result_data:
            filtered = {}
            for key in selected:
                if key == "vendor_name":
                    vendor = self.result_data.get("vendor")
                    if isinstance(vendor, dict):
                        filtered["vendor_name"] = vendor.get("name")
                    else:
                        filtered["vendor_name"] = self.result_data.get("vendor_name")
                elif key == "net_amount":
                    filtered["net_amount"] = self.result_data.get("subtotal") or self.result_data.get("net_amount")
                else:
                    filtered[key] = self.result_data.get(key)
            display_data = filtered
        else:
            display_data = self.result_data

        json_str = json.dumps(display_data, indent=2, sort_keys=True)
        self._log(json_str)

    def _log(self, msg):
        self.textbox.configure(state="normal")
        self.textbox.insert("end", msg + "\n")
        self.textbox.see("end")
        self.textbox.configure(state="disabled")

    def _clear_page(self):
        self.textbox.configure(state="normal")
        self.textbox.delete("1.0", "end")
        self.textbox.configure(state="disabled")
        self.result_data = None
        self.save_btn.configure(state="disabled")
        self.copy_btn.configure(state="disabled")
        self.status_label.configure(text="Cleared.", fg=TEXT_DIM)

    def copy_json_to_clipboard(self):
        if self.result_data:
            self.clipboard_clear()
            self.clipboard_append(json.dumps(self.result_data, indent=4))
            messagebox.showinfo("Success", "JSON Data copied to clipboard!")

    def save_excel(self):
        if not self.result_data:
            return
        inv_num = self.result_data.get("invoice_number") or "invoice"
        default_name = f"invoice_{re.sub(r'[^a-zA-Z0-9]', '_', str(inv_num))}.xlsx"
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=default_name,
            title="Save Invoice Excel")
        if path:
            try:
                build_excel(path, self.result_data)
                messagebox.showinfo("Success", f"Excel report saved to:\n{path}")
                self._log(f"Saved Excel to '{path}'")
            except Exception as e:
                messagebox.showerror("Error Saving Excel", str(e))

    def reset_ui_state(self):
        self.extract_btn.configure(state="normal", text="⚡  Extract")
        self.status_label.configure(text="Extraction complete.", fg="#4ecb8d")

    def _show_help(self):
        help_win = tk.Toplevel(self)
        help_win.title("Help — Invoice Extractor")
        help_win.configure(bg=BG_DARK)
        help_win.geometry("500x400")
        help_win.minsize(400, 300)
        help_text = tk.Text(help_win, bg=LOGGER_BG, fg=LOGGER_FG, font=("Consolas", 11),
                            wrap="word", bd=0, padx=14, pady=14)
        help_text.pack(fill="both", expand=True, padx=10, pady=10)
        help_text.insert("1.0",
            "Invoice Extractor — Help\n"
            "═══════════════════════════\n\n"
            "1. Enter your Groq API key in the right panel.\n\n"
            "2. Open a PDF or image invoice using the 📂 button\n"
            "   or open a directory with 📁.\n\n"
            "3. Select the fields you want to extract using the\n"
            "   checkboxes.\n\n"
            "4. Click '⚡ Extract' to analyze the current page.\n\n"
            "5. Results appear in the logger below. You can:\n"
            "   • Copy the JSON with 📋\n"
            "   • Save to Excel with 💾\n\n"
            "Navigation:\n"
            "  ⏮ ◀  Page navigation  ▶ ⏭\n"
            "  🔍+ / 🔍−  Zoom in / out\n"
            "  ⊞  Fit to screen\n"
        )
        help_text.configure(state="disabled")

if __name__ == "__main__":
    app = InvoiceExtractorApp()
    app.mainloop()