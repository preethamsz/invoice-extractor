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

class InvoiceExtractorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Invoice Extractor")
        self.geometry("1280x800")
        self.minsize(1024, 700)

        # State variables
        self.file_path = None
        self.pdf_images = []       # list of base64 strings
        self.pdf_pil_images = []   # list of PIL Image objects for display
        self.current_page = 0
        self.total_pages = 0
        self.result_data = None
        self.zoom_level = 1.0
        self.field_vars = {}
        self.displayed_image = None  # keep reference

        self.build_ui()

    # ────────────────── UI Construction ──────────────────
    def build_ui(self):
        self.grid_columnconfigure(0, weight=3)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(0, weight=1)

        # ═══ LEFT PANEL — PDF / Image Viewer ═══
        self.left_panel = ctk.CTkFrame(self, corner_radius=0, fg_color="#2b2b2b")
        self.left_panel.grid(row=0, column=0, sticky="nsew")
        self.left_panel.grid_rowconfigure(1, weight=1)
        self.left_panel.grid_columnconfigure(0, weight=1)

        # — Top toolbar —
        self.toolbar = ctk.CTkFrame(self.left_panel, height=40, fg_color="#3c3c3c", corner_radius=0)
        self.toolbar.grid(row=0, column=0, sticky="ew")
        self.toolbar.grid_columnconfigure(1, weight=1)

        # Navigation buttons
        nav_frame = ctk.CTkFrame(self.toolbar, fg_color="transparent")
        nav_frame.grid(row=0, column=0, padx=8, pady=4)

        self.btn_first = ctk.CTkButton(nav_frame, text="|◀", width=30, height=28, fg_color="#555", hover_color="#777", command=self.first_page)
        self.btn_first.pack(side="left", padx=2)
        self.btn_prev = ctk.CTkButton(nav_frame, text="◀", width=30, height=28, fg_color="#555", hover_color="#777", command=self.prev_page)
        self.btn_prev.pack(side="left", padx=2)

        self.page_label = ctk.CTkLabel(nav_frame, text="Page 0 of 0", font=ctk.CTkFont(size=12), text_color="#ddd")
        self.page_label.pack(side="left", padx=8)

        self.btn_next = ctk.CTkButton(nav_frame, text="▶", width=30, height=28, fg_color="#555", hover_color="#777", command=self.next_page)
        self.btn_next.pack(side="left", padx=2)
        self.btn_last = ctk.CTkButton(nav_frame, text="▶|", width=30, height=28, fg_color="#555", hover_color="#777", command=self.last_page)
        self.btn_last.pack(side="left", padx=2)

        # Zoom controls
        zoom_frame = ctk.CTkFrame(self.toolbar, fg_color="transparent")
        zoom_frame.grid(row=0, column=1, padx=8, pady=4, sticky="e")

        self.btn_zoom_in = ctk.CTkButton(zoom_frame, text="+", width=30, height=28, fg_color="#555", hover_color="#777", command=self.zoom_in)
        self.btn_zoom_in.pack(side="left", padx=2)
        self.btn_zoom_out = ctk.CTkButton(zoom_frame, text="−", width=30, height=28, fg_color="#555", hover_color="#777", command=self.zoom_out)
        self.btn_zoom_out.pack(side="left", padx=2)
        self.zoom_label = ctk.CTkLabel(zoom_frame, text="100%", font=ctk.CTkFont(size=12), text_color="#ddd")
        self.zoom_label.pack(side="left", padx=8)

        # — Canvas for document display —
        self.canvas_frame = ctk.CTkFrame(self.left_panel, fg_color="#4a4a4a", corner_radius=0)
        self.canvas_frame.grid(row=1, column=0, sticky="nsew")
        self.canvas_frame.grid_rowconfigure(0, weight=1)
        self.canvas_frame.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(self.canvas_frame, bg="#3a3a3a", highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")

        self.v_scroll = tk.Scrollbar(self.canvas_frame, orient="vertical", command=self.canvas.yview)
        self.v_scroll.grid(row=0, column=1, sticky="ns")
        self.h_scroll = tk.Scrollbar(self.canvas_frame, orient="horizontal", command=self.canvas.xview)
        self.h_scroll.grid(row=1, column=0, sticky="ew")
        self.canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)

        # Mouse wheel scroll
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        # Auto-fit on resize
        self.canvas.bind("<Configure>", self._on_canvas_resize)

        # Page indicator at bottom-left
        self.bottom_bar = ctk.CTkFrame(self.left_panel, height=28, fg_color="#3c3c3c", corner_radius=0)
        self.bottom_bar.grid(row=2, column=0, sticky="ew")
        self.page_indicator = ctk.CTkLabel(self.bottom_bar, text="", font=ctk.CTkFont(size=11), text_color="#aaa")
        self.page_indicator.pack(side="left", padx=10)

        # ═══ RIGHT PANEL — Controls & Results ═══
        self.right_panel = ctk.CTkFrame(self, corner_radius=0, fg_color="#1e1e2e")
        self.right_panel.grid(row=0, column=1, sticky="nsew")
        self.right_panel.grid_rowconfigure(4, weight=1)
        self.right_panel.grid_columnconfigure(0, weight=1)

        # — Branding header —
        header_frame = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        ctk.CTkLabel(header_frame, text="Invoice Extractor", font=ctk.CTkFont(size=28, weight="bold"), text_color="#e0a050").pack(side="left", padx=(0, 10))

        # — API Key (compact) —
        api_frame = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        api_frame.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 10))
        ctk.CTkLabel(api_frame, text="API Key:", font=ctk.CTkFont(size=12, weight="bold"), text_color="#ccc").pack(side="left", padx=(0, 8))
        self.api_entry = ctk.CTkEntry(api_frame, show="*", placeholder_text="gsk_...", height=32, border_color="#4a4a6a")
        self.api_entry.pack(side="left", fill="x", expand=True)

        # — Field checkboxes —
        field_frame = ctk.CTkFrame(self.right_panel, fg_color="#292940", corner_radius=8)
        field_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 10))
        ctk.CTkLabel(field_frame, text="Field:", font=ctk.CTkFont(size=13, weight="bold"), text_color="#ccc").grid(row=0, column=0, columnspan=4, sticky="w", padx=12, pady=(10, 5))

        for i, (key, label) in enumerate(FIELD_OPTIONS):
            var = ctk.BooleanVar(value=True)
            self.field_vars[key] = var
            row_idx = 1 + i // 2
            col_idx = (i % 2) * 2
            cb = ctk.CTkCheckBox(field_frame, text=label, variable=var, font=ctk.CTkFont(size=12),
                                 fg_color="#2196F3", hover_color="#42A5F5", border_color="#666",
                                 text_color="#ddd", checkbox_width=20, checkbox_height=20)
            cb.grid(row=row_idx, column=col_idx, columnspan=2, sticky="w", padx=12, pady=4)

        field_frame.grid_columnconfigure(0, weight=1)
        field_frame.grid_columnconfigure(2, weight=1)

        # — Buttons row: Browse + Extract —
        btn_row = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        btn_row.grid(row=3, column=0, sticky="ew", padx=20, pady=(0, 10))
        btn_row.grid_columnconfigure(0, weight=1)
        btn_row.grid_columnconfigure(1, weight=1)

        self.browse_btn = ctk.CTkButton(btn_row, text="Open File", fg_color="#555", hover_color="#777",
                                        height=40, font=ctk.CTkFont(size=14), command=self.browse_file)
        self.browse_btn.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        self.extract_btn = ctk.CTkButton(btn_row, text="Extract", fg_color="#00897B", hover_color="#00AB9A",
                                         height=40, font=ctk.CTkFont(size=14, weight="bold"), command=self.start_extraction)
        self.extract_btn.grid(row=0, column=1, sticky="ew", padx=(5, 0))

        self.status_label = ctk.CTkLabel(self.right_panel, text="", text_color="gray", font=ctk.CTkFont(size=11))
        self.status_label.grid(row=3, column=0, sticky="w", padx=20, pady=(42, 0))

        # — Results text area —
        self.textbox = ctk.CTkTextbox(self.right_panel, font=ctk.CTkFont(family="Consolas", size=13),
                                      fg_color="#0d1117", text_color="#58d1c9", wrap="word", corner_radius=6,
                                      border_color="#333", border_width=1)
        self.textbox.grid(row=4, column=0, sticky="nsew", padx=20, pady=(5, 10))
        self.textbox.insert("0.0", "Open an invoice PDF or image, then click Extract.\nResults will appear here.")
        self.textbox.configure(state="disabled")

        # — Bottom action buttons —
        bottom_btns = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        bottom_btns.grid(row=5, column=0, sticky="ew", padx=20, pady=(0, 20))
        bottom_btns.grid_columnconfigure(0, weight=1)
        bottom_btns.grid_columnconfigure(1, weight=1)

        self.save_btn = ctk.CTkButton(bottom_btns, text="Save Information", fg_color="#555", hover_color="#777",
                                      height=40, font=ctk.CTkFont(size=13), command=self.save_excel, state="disabled")
        self.save_btn.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        self.copy_btn = ctk.CTkButton(bottom_btns, text="Copy JSON", fg_color="#00897B", hover_color="#00AB9A",
                                      height=40, font=ctk.CTkFont(size=13), command=self.copy_json_to_clipboard, state="disabled")
        self.copy_btn.grid(row=0, column=1, sticky="ew", padx=(5, 0))

    # ────────────────── PDF Viewer Logic ──────────────────
    def _fit_zoom_to_canvas(self):
        """Calculate zoom level to fit image width to canvas."""
        if not self.pdf_pil_images:
            return
        img = self.pdf_pil_images[self.current_page]
        canvas_w = self.canvas.winfo_width()
        if canvas_w > 1 and img.width > 0:
            self.zoom_level = (canvas_w - 20) / img.width  # 20px padding

    def _on_canvas_resize(self, event=None):
        """Re-fit image when canvas is resized (only if auto-fit)."""
        if self.pdf_pil_images and not hasattr(self, '_manual_zoom'):
            self._fit_zoom_to_canvas()
            self._render_current_page()

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def show_page(self):
        if not self.pdf_pil_images:
            return
        # Auto-fit on first load
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
        # Center image horizontally
        x_offset = max((canvas_w - w) // 2, 0)

        self.canvas.delete("all")
        # White page background behind the image
        self.canvas.create_rectangle(x_offset, 10, x_offset + w, 10 + h, fill="white", outline="#555")
        self.canvas.create_image(x_offset, 10, anchor="nw", image=self.displayed_image)
        self.canvas.configure(scrollregion=(0, 0, max(canvas_w, w + 20), h + 20))

        self.page_label.configure(text=f"Page {self.current_page + 1} of {self.total_pages}")
        self.page_indicator.configure(text=f"{self.current_page + 1} of {self.total_pages}")
        self.zoom_label.configure(text=f"{int(self.zoom_level * 100)}%")

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

        self.file_path = path
        ext = os.path.splitext(path)[1].lower()

        if ext == ".pdf":
            self.status_label.configure(text="Converting PDF...", text_color="#ff8c69")
            threading.Thread(target=self.process_pdf, args=(path,), daemon=True).start()
        else:
            # Single image
            self.pdf_images = []
            img = Image.open(path)
            self.pdf_pil_images = [img]
            self.total_pages = 1
            self.current_page = 0
            with open(path, "rb") as f:
                self.pdf_images = [base64.b64encode(f.read()).decode()]
            self.zoom_level = 1.0
            self.show_page()
            self.status_label.configure(text=f"Loaded: {os.path.basename(path)}", text_color="#4ecb8d")

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
                text=f"Ready: {self.total_pages} page(s)", text_color="#4ecb8d"))
        except Exception as e:
            self.after(0, lambda: self.status_label.configure(text="Error converting PDF", text_color="#e63946"))
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

        self.extract_btn.configure(state="disabled", text="Processing...")
        self.save_btn.configure(state="disabled")
        self.copy_btn.configure(state="disabled")
        self.status_label.configure(text="Analyzing via Groq AI...", text_color="white")

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

        json_str = json.dumps(display_data, indent=4)
        self.textbox.configure(state="normal")
        self.textbox.delete("0.0", "end")
        self.textbox.insert("0.0", json_str)
        self.textbox.configure(state="disabled")

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
            except Exception as e:
                messagebox.showerror("Error Saving Excel", str(e))

    def reset_ui_state(self):
        self.extract_btn.configure(state="normal", text="Extract")
        self.status_label.configure(text="Extraction complete.", text_color="#4ecb8d")

if __name__ == "__main__":
    app = InvoiceExtractorApp()
    app.mainloop()