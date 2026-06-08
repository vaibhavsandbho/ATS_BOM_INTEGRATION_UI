"""
ATS BOM Automation Tool v1.0
Convert Vault CSV exports into Clarity ERP Item Import and BOM Import Excel files.

Themed to match the official ATS Group website (https://www.ats-group.com)
Brand Colors: ATS Blue #0e4092, ATS Gold #ffcd06
Brand Fonts: Exo (headings), Open Sans (body)
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk
from PIL import Image
import os
import sys
import csv
import threading
import time
import importlib.util
from datetime import datetime


# ═══════════════════════════════════════════════════════════════════════════════
#  EXTERNAL SCRIPT PATHS (Quentin's Vault-to-Clarity converters)
# ═══════════════════════════════════════════════════════════════════════════════

SCRIPT_ITEM_IMPORT = r"C:\Users\vaibhavs\OneDrive - ATS Group\Desktop\Item Vault to clarity.py"
SCRIPT_BOM_IMPORT  = r"C:\Users\vaibhavs\OneDrive - ATS Group\Desktop\Vault to clarity Claude.py"


# ═══════════════════════════════════════════════════════════════════════════════
#  ATS GROUP BRAND DESIGN TOKENS  (from https://www.ats-group.com)
# ═══════════════════════════════════════════════════════════════════════════════

ATS = {
    # ── Primary Brand ──────────────────────────────────────────────────────
    "blue":              "#0e4092",      # Primary ATS Blue
    "blue_dark":         "#072049",      # Hover/active state
    "blue_mid":          "#0b3375",      # Intermediate
    "blue_light":        "#cfe0fa",      # Light tint
    "blue_surface":      "#e8eff9",      # Card/surface tint

    # ── Secondary Brand ────────────────────────────────────────────────────
    "gold":              "#ffcd06",      # ATS Gold/Yellow
    "gold_dark":         "#d1a700",      # Gold hover
    "gold_light":        "#fff8da",      # Gold tint

    # ── Neutrals ───────────────────────────────────────────────────────────
    "white":             "#ffffff",
    "bg":                "#f5f7fa",      # Window background
    "card":              "#ffffff",      # Card background
    "input_bg":          "#f5f5f5",      # Input field bg
    "border":            "#e0e4ec",      # General borders
    "border_light":      "#eef1f6",      # Subtle borders
    "divider":           "#cacaca",      # Divider lines

    # ── Text ───────────────────────────────────────────────────────────────
    "text_dark":         "#1a1a2e",      # Headings
    "text":              "#333333",      # Body text (ATS site uses #333)
    "text_muted":        "#8a8a8a",      # Muted/caption
    "text_on_blue":      "#ffffff",      # Text on blue
    "text_on_blue_sub":  "#a8c4e8",      # Subtle on blue
    "text_on_dark":      "#ffffff",
    "text_on_dark_sub":  "#b0bec5",

    # ── Status ─────────────────────────────────────────────────────────────
    "success":           "#3adb76",      # ATS success green
    "success_dark":      "#22bb5b",
    "success_bg":        "#e1faea",
    "warning":           "#ffae00",      # ATS warning orange
    "warning_dark":      "#cc8b00",
    "warning_bg":        "#fff3d9",
    "error":             "#cc4b37",      # ATS alert red
    "error_dark":        "#a53b2a",
    "error_bg":          "#f7e4e1",

    # ── Footer ─────────────────────────────────────────────────────────────
    "footer_bg":         "#1a1a2e",
    "footer_text":       "#a8b2c1",
}

# Fonts: ATS uses Exo for headings and Open Sans for body.
# Fallback to Segoe UI on Windows if web fonts not installed.
HEADING_FONT = "Segoe UI Semibold"  # Best Windows fallback for Exo
BODY_FONT = "Segoe UI"             # Best Windows fallback for Open Sans

FONTS = {
    "app_title":      (HEADING_FONT, 20),
    "app_subtitle":   (BODY_FONT, 11),
    "section_title":  (HEADING_FONT, 14),
    "section_icon":   (BODY_FONT, 14),
    "body":           (BODY_FONT, 12),
    "body_small":     (BODY_FONT, 11),
    "body_bold":      ("Segoe UI Semibold", 12),
    "caption":        (BODY_FONT, 10),
    "button":         ("Segoe UI Semibold", 12),
    "button_lg":      ("Segoe UI Semibold", 14),
    "footer":         (BODY_FONT, 10),
    "badge":          ("Segoe UI Semibold", 10),
    "table_header":   ("Segoe UI Semibold", 11),
    "table_cell":     (BODY_FONT, 11),
    "dropzone_icon":  ("Segoe UI Light", 36),
    "dropzone_text":  (BODY_FONT, 13),
    "progress_pct":   ("Segoe UI Semibold", 13),
    "status_msg":     (BODY_FONT, 11),
    "version_badge":  ("Segoe UI Semibold", 10),
}


# ═══════════════════════════════════════════════════════════════════════════════
#  CARD FRAME
# ═══════════════════════════════════════════════════════════════════════════════
class CardFrame(ctk.CTkFrame):
    """ATS-styled card container with optional titled header and gold accent."""

    def __init__(self, master, title="", icon="", **kwargs):
        super().__init__(
            master,
            fg_color=ATS["card"],
            corner_radius=8,
            border_width=1,
            border_color=ATS["border_light"],
            **kwargs,
        )

        if title:
            # Gold top-accent line
            accent = ctk.CTkFrame(self, fg_color=ATS["gold"], height=3, corner_radius=0)
            accent.pack(fill="x", padx=0, pady=0)

            header = ctk.CTkFrame(self, fg_color="transparent", height=40)
            header.pack(fill="x", padx=20, pady=(14, 4))

            if icon:
                icon_lbl = ctk.CTkLabel(
                    header, text=icon, font=FONTS["section_icon"],
                    text_color=ATS["blue"],
                )
                icon_lbl.pack(side="left", padx=(0, 8))

            title_lbl = ctk.CTkLabel(
                header, text=title.upper(),
                font=FONTS["section_title"],
                text_color=ATS["blue"],
                anchor="w",
            )
            title_lbl.pack(side="left", fill="x", expand=True)

            sep = ctk.CTkFrame(self, fg_color=ATS["border"], height=1)
            sep.pack(fill="x", padx=20, pady=(4, 0))


# ═══════════════════════════════════════════════════════════════════════════════
#  VALIDATION ITEM
# ═══════════════════════════════════════════════════════════════════════════════
class ValidationItem(ctk.CTkFrame):
    """Single validation row with ATS-styled status indicator."""

    def __init__(self, master, text, **kwargs):
        super().__init__(master, fg_color="transparent", height=32, **kwargs)

        self.icon_label = ctk.CTkLabel(
            self, text="●", font=(BODY_FONT, 12),
            text_color=ATS["divider"], width=24,
        )
        self.icon_label.pack(side="left", padx=(0, 8))

        self.text_label = ctk.CTkLabel(
            self, text=text, font=FONTS["body"],
            text_color=ATS["text_muted"], anchor="w",
        )
        self.text_label.pack(side="left", fill="x", expand=True)

        self.value_label = ctk.CTkLabel(
            self, text="—", font=FONTS["body_bold"],
            text_color=ATS["divider"], anchor="e",
        )
        self.value_label.pack(side="right", padx=(8, 0))

    def set_success(self, value=""):
        self.icon_label.configure(text="✓", text_color=ATS["success_dark"])
        self.text_label.configure(text_color=ATS["text"])
        if value:
            self.value_label.configure(text=value, text_color=ATS["success_dark"])

    def set_error(self, value=""):
        self.icon_label.configure(text="✕", text_color=ATS["error"])
        self.text_label.configure(text_color=ATS["error"])
        if value:
            self.value_label.configure(text=value, text_color=ATS["error"])

    def set_pending(self):
        self.icon_label.configure(text="●", text_color=ATS["divider"])
        self.text_label.configure(text_color=ATS["text_muted"])
        self.value_label.configure(text="—", text_color=ATS["divider"])


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ═══════════════════════════════════════════════════════════════════════════════
class ATSBOMAutomationTool(ctk.CTk):

    REQUIRED_COLUMNS = [
        "Part Number", "Description", "Material", "Category",
        "Revision", "State", "Unit of Measure",
    ]

    def __init__(self):
        super().__init__()

        self.title("ATS BOM Automation Tool — v1.0")
        self.geometry("1200x920")
        self.minsize(1100, 800)
        self.configure(fg_color=ATS["bg"])

        # Center on screen
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"1200x920+{(sw-1200)//2}+{(sh-920)//2}")

        # State
        self.csv_filepath = None
        self.csv_data = []
        self.csv_columns = []
        self.output_dir = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))
        self.gen_item_import = tk.BooleanVar(value=True)
        self.gen_bom_import = tk.BooleanVar(value=True)
        self.is_processing = False
        self.error_log = []

        self._build_header()
        self._build_main_area()
        self._build_footer()
        self._load_logo()

    # ═══════════════════════════════════════════════════════════════════════
    #  HEADER — ATS Blue with gold accent bottom line
    # ═══════════════════════════════════════════════════════════════════════
    def _build_header(self):
        # Gold accent strip at the very top
        gold_strip = ctk.CTkFrame(self, fg_color=ATS["gold"], corner_radius=0, height=4)
        gold_strip.pack(fill="x", side="top")

        self.header = ctk.CTkFrame(
            self, fg_color=ATS["blue"], corner_radius=0, height=72,
        )
        self.header.pack(fill="x", side="top")
        self.header.pack_propagate(False)

        inner = ctk.CTkFrame(self.header, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=28, pady=0)

        # Left: Logo + Title
        left = ctk.CTkFrame(inner, fg_color="transparent")
        left.pack(side="left", fill="y")

        self.logo_label = ctk.CTkLabel(left, text="", width=48, height=48)
        self.logo_label.pack(side="left", padx=(0, 16), pady=12)

        title_block = ctk.CTkFrame(left, fg_color="transparent")
        title_block.pack(side="left", fill="y", pady=12)

        title = ctk.CTkLabel(
            title_block, text="ATS BOM Automation Tool",
            font=FONTS["app_title"], text_color=ATS["text_on_blue"],
        )
        title.pack(anchor="w")

        subtitle = ctk.CTkLabel(
            title_block, text="Vault CSV  →  Clarity ERP Import Files",
            font=FONTS["app_subtitle"], text_color=ATS["text_on_blue_sub"],
        )
        subtitle.pack(anchor="w")

        # Right: Refresh button + version badge
        right = ctk.CTkFrame(inner, fg_color="transparent")
        right.pack(side="right", fill="y")

        ver_badge = ctk.CTkFrame(right, fg_color=ATS["gold"], corner_radius=4)
        ver_badge.pack(side="right", pady=22)

        ver_label = ctk.CTkLabel(
            ver_badge, text="  v1.0  ",
            font=FONTS["version_badge"], text_color=ATS["blue_dark"],
        )
        ver_label.pack(padx=10, pady=4)

        # Refresh button
        refresh_btn = ctk.CTkButton(
            right, text="  🔄  Refresh  ", font=FONTS["button"],
            fg_color="#1565C0", hover_color=ATS["blue_dark"],
            text_color=ATS["text_on_blue"], corner_radius=6,
            height=32, width=110, command=self._refresh_all,
        )
        refresh_btn.pack(side="right", padx=(0, 14), pady=22)

    def _load_logo(self):
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ats_logo.png")
        if os.path.exists(logo_path):
            try:
                logo_img = ctk.CTkImage(
                    light_image=Image.open(logo_path),
                    dark_image=Image.open(logo_path),
                    size=(120, 48),
                )
                self.logo_label.configure(image=logo_img, text="", width=120)
                self.logo_label._image = logo_img
            except Exception:
                self._set_text_logo()
        else:
            self._set_text_logo()

    def _set_text_logo(self):
        self.logo_label.configure(
            text="ATS", font=("Segoe UI Bold", 20),
            text_color=ATS["gold"],
        )

    # ═══════════════════════════════════════════════════════════════════════
    #  MAIN CONTENT
    # ═══════════════════════════════════════════════════════════════════════
    def _build_main_area(self):
        self.main_scroll = ctk.CTkScrollableFrame(
            self, fg_color=ATS["bg"], corner_radius=0,
            scrollbar_button_color=ATS["border"],
            scrollbar_button_hover_color=ATS["blue"],
        )
        self.main_scroll.pack(fill="both", expand=True)

        content = ctk.CTkFrame(self.main_scroll, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=28, pady=20)

        # Row 1: File Upload + Validation
        row1 = ctk.CTkFrame(content, fg_color="transparent")
        row1.pack(fill="x", pady=(0, 14))
        row1.columnconfigure(0, weight=3)
        row1.columnconfigure(1, weight=2)
        self._build_file_upload(row1)
        self._build_validation_panel(row1)

        # Row 2: Options + Output Folder
        row2 = ctk.CTkFrame(content, fg_color="transparent")
        row2.pack(fill="x", pady=(0, 14))
        row2.columnconfigure(0, weight=1)
        row2.columnconfigure(1, weight=2)
        self._build_generation_options(row2)
        self._build_output_folder(row2)

        # Row 3: Processing
        self._build_processing_area(content)

        # Row 4: Generated Files
        self._build_generated_files(content)

        # Row 5: Error Log
        self._build_error_log(content)

    # ─── 1. FILE UPLOAD ──────────────────────────────────────────────────
    def _build_file_upload(self, parent):
        card = CardFrame(parent, title="Select Vault CSV File", icon="📂")
        card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=(12, 20))

        # Dropzone — ATS blue dashed style
        self.dropzone = ctk.CTkFrame(
            body, fg_color=ATS["blue_surface"], corner_radius=8,
            border_width=2, border_color=ATS["blue_light"], height=130,
        )
        self.dropzone.pack(fill="x", pady=(0, 12))
        self.dropzone.pack_propagate(False)

        dz_inner = ctk.CTkFrame(self.dropzone, fg_color="transparent")
        dz_inner.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(
            dz_inner, text="⬆", font=FONTS["dropzone_icon"],
            text_color=ATS["blue"],
        ).pack()

        ctk.CTkLabel(
            dz_inner, text="Drag & Drop your Vault CSV file here",
            font=FONTS["dropzone_text"], text_color=ATS["text"],
        ).pack(pady=(4, 2))

        ctk.CTkLabel(
            dz_inner, text="or click Browse to select a file",
            font=FONTS["caption"], text_color=ATS["text_muted"],
        ).pack()

        # Bottom: Browse + filename
        bottom = ctk.CTkFrame(body, fg_color="transparent")
        bottom.pack(fill="x")

        ctk.CTkButton(
            bottom, text="  Browse Files  ", font=FONTS["button"],
            fg_color=ATS["blue"], hover_color=ATS["blue_dark"],
            corner_radius=6, height=36, command=self._browse_csv,
        ).pack(side="left")

        self.file_display = ctk.CTkLabel(
            bottom, text="No file selected",
            font=FONTS["body_small"], text_color=ATS["text_muted"], anchor="w",
        )
        self.file_display.pack(side="left", padx=(14, 0), fill="x", expand=True)

    # ─── 2. VALIDATION PANEL ────────────────────────────────────────────
    def _build_validation_panel(self, parent):
        card = CardFrame(parent, title="Validation Status", icon="✓")
        card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=(12, 20))

        self.val_csv_loaded = ValidationItem(body, "CSV File Loaded")
        self.val_csv_loaded.pack(fill="x", pady=3)

        self.val_columns = ValidationItem(body, "Required Columns Found")
        self.val_columns.pack(fill="x", pady=3)

        self.val_records = ValidationItem(body, "Total Records")
        self.val_records.pack(fill="x", pady=3)

        self.val_overall = ValidationItem(body, "Validation Status")
        self.val_overall.pack(fill="x", pady=3)

        # Status banner
        self.val_banner = ctk.CTkFrame(
            body, fg_color=ATS["input_bg"], corner_radius=6, height=36,
        )
        self.val_banner.pack(fill="x", pady=(10, 0))

        self.val_banner_label = ctk.CTkLabel(
            self.val_banner, text="Awaiting CSV file...",
            font=FONTS["body_small"], text_color=ATS["text_muted"],
        )
        self.val_banner_label.pack(pady=8)

    # ─── 3. GENERATION OPTIONS ──────────────────────────────────────────
    def _build_generation_options(self, parent):
        card = CardFrame(parent, title="Generation Options", icon="⚙")
        card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=(12, 20))

        # Item Import checkbox — linked to Item Vault to clarity.py
        cb1 = ctk.CTkCheckBox(
            body, text="  Generate Item Import File  —  Item Vault to clarity.py",
            font=FONTS["body"], text_color=ATS["text"],
            variable=self.gen_item_import,
            fg_color=ATS["blue"], hover_color=ATS["blue_dark"],
            border_color=ATS["divider"], corner_radius=4,
            border_width=2, checkmark_color="white",
        )
        cb1.pack(anchor="w", pady=(6, 2))

        ctk.CTkLabel(
            body, text=f"📄 {SCRIPT_ITEM_IMPORT}",
            font=FONTS["caption"], text_color=ATS["text_muted"], anchor="w",
        ).pack(anchor="w", padx=(28, 0), pady=(0, 10))

        # BOM Import checkbox — linked to Vault to clarity Claude.py
        cb2 = ctk.CTkCheckBox(
            body, text="  Generate BOM Import File  —  Vault to clarity Claude.py",
            font=FONTS["body"], text_color=ATS["text"],
            variable=self.gen_bom_import,
            fg_color=ATS["blue"], hover_color=ATS["blue_dark"],
            border_color=ATS["divider"], corner_radius=4,
            border_width=2, checkmark_color="white",
        )
        cb2.pack(anchor="w", pady=(6, 2))

        ctk.CTkLabel(
            body, text=f"📄 {SCRIPT_BOM_IMPORT}",
            font=FONTS["caption"], text_color=ATS["text_muted"], anchor="w",
        ).pack(anchor="w", padx=(28, 0), pady=(0, 4))

    # ─── 4. OUTPUT FOLDER ───────────────────────────────────────────────
    def _build_output_folder(self, parent):
        card = CardFrame(parent, title="Output Folder", icon="📁")
        card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=(12, 20))

        row = ctk.CTkFrame(body, fg_color="transparent")
        row.pack(fill="x", pady=(4, 0))

        ctk.CTkEntry(
            row, textvariable=self.output_dir, font=FONTS["body"],
            fg_color=ATS["input_bg"], border_color=ATS["divider"],
            border_width=1, corner_radius=6, text_color=ATS["text"], height=36,
        ).pack(side="left", fill="x", expand=True, padx=(0, 10))

        ctk.CTkButton(
            row, text="  Browse  ", font=FONTS["button"],
            fg_color=ATS["blue"], hover_color=ATS["blue_dark"],
            corner_radius=6, height=36, command=self._browse_output,
        ).pack(side="right")

    # ─── 5. PROCESSING AREA ────────────────────────────────────────────
    def _build_processing_area(self, parent):
        card = CardFrame(parent, title="Processing", icon="▶")
        card.pack(fill="x", pady=(0, 14))

        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="x", padx=20, pady=(12, 20))

        top = ctk.CTkFrame(body, fg_color="transparent")
        top.pack(fill="x", pady=(0, 14))

        # ATS Gold-accented process button
        self.process_btn = ctk.CTkButton(
            top, text="  ▶  Start Processing  ", font=FONTS["button_lg"],
            fg_color=ATS["gold"], hover_color=ATS["gold_dark"],
            text_color=ATS["blue_dark"], corner_radius=6, height=44,
            command=self._start_processing,
        )
        self.process_btn.pack(side="left")

        self.status_label = ctk.CTkLabel(
            top, text="Ready to process", font=FONTS["status_msg"],
            text_color=ATS["text_muted"], anchor="w",
        )
        self.status_label.pack(side="left", padx=(18, 0), fill="x", expand=True)

        self.pct_label = ctk.CTkLabel(
            top, text="0%", font=FONTS["progress_pct"], text_color=ATS["blue"],
        )
        self.pct_label.pack(side="right")

        # Progress bar — ATS blue on light track
        self.progress_bar = ctk.CTkProgressBar(
            body, fg_color=ATS["border"], progress_color=ATS["blue"],
            corner_radius=5, height=10,
        )
        self.progress_bar.pack(fill="x")
        self.progress_bar.set(0)

    # ─── 6. GENERATED FILES ────────────────────────────────────────────
    def _build_generated_files(self, parent):
        card = CardFrame(parent, title="Generated Files", icon="📄")
        card.pack(fill="x", pady=(0, 14))

        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="x", padx=20, pady=(12, 20))

        files_frame = ctk.CTkFrame(body, fg_color="transparent")
        files_frame.pack(fill="x", pady=(0, 12))
        files_frame.columnconfigure(0, weight=1)
        files_frame.columnconfigure(1, weight=1)

        self.item_file_card = self._create_file_card(files_frame, "Item Import", "ItemImport.xlsx", 0)
        self.bom_file_card  = self._create_file_card(files_frame, "BOM Import",  "BOMImport.xlsx",  1)

        btn_row = ctk.CTkFrame(body, fg_color="transparent")
        btn_row.pack(fill="x")

        ctk.CTkButton(
            btn_row, text="  📂  Open Output Folder  ", font=FONTS["button"],
            fg_color=ATS["input_bg"], hover_color=ATS["border"],
            text_color=ATS["text"], border_width=1, border_color=ATS["divider"],
            corner_radius=6, height=36, command=self._open_output_folder,
        ).pack(side="left")

    def _create_file_card(self, parent, title, filename, col):
        frame = ctk.CTkFrame(
            parent, fg_color=ATS["blue_surface"], corner_radius=8,
            border_width=1, border_color=ATS["blue_light"],
        )
        frame.grid(
            row=0, column=col, sticky="nsew",
            padx=(0 if col == 0 else 6, 6 if col == 0 else 0),
        )

        inner = ctk.CTkFrame(frame, fg_color="transparent")
        inner.pack(fill="x", padx=16, pady=14)

        ctk.CTkLabel(
            inner, text="📊", font=("Segoe UI", 24),
        ).pack(side="left", padx=(0, 12))

        info = ctk.CTkFrame(inner, fg_color="transparent")
        info.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(
            info, text=title, font=FONTS["body_bold"],
            text_color=ATS["blue"], anchor="w",
        ).pack(anchor="w")

        f_label = ctk.CTkLabel(
            info, text=filename, font=FONTS["caption"],
            text_color=ATS["text_muted"], anchor="w",
        )
        f_label.pack(anchor="w")

        badge = ctk.CTkFrame(inner, fg_color=ATS["border_light"], corner_radius=4)
        badge.pack(side="right")

        badge_label = ctk.CTkLabel(
            badge, text="  Pending  ", font=FONTS["badge"],
            text_color=ATS["text_muted"],
        )
        badge_label.pack(padx=6, pady=3)

        return {"frame": frame, "badge": badge, "badge_label": badge_label, "filename_label": f_label}

    def _update_file_card_status(self, card, status, filename=""):
        if status == "success":
            card["badge"].configure(fg_color=ATS["success_bg"])
            card["badge_label"].configure(text="  ✓ Generated  ", text_color=ATS["success_dark"])
            if filename:
                card["filename_label"].configure(text=filename)
        elif status == "error":
            card["badge"].configure(fg_color=ATS["error_bg"])
            card["badge_label"].configure(text="  ✕ Failed  ", text_color=ATS["error"])
        elif status == "skipped":
            card["badge"].configure(fg_color=ATS["warning_bg"])
            card["badge_label"].configure(text="  — Skipped  ", text_color=ATS["warning_dark"])
        else:
            card["badge"].configure(fg_color=ATS["border_light"])
            card["badge_label"].configure(text="  Pending  ", text_color=ATS["text_muted"])

    # ─── 7. ERROR LOG ──────────────────────────────────────────────────
    def _build_error_log(self, parent):
        card = CardFrame(parent, title="Error Log", icon="⚠")
        card.pack(fill="x", pady=(0, 8))

        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="x", padx=20, pady=(12, 20))

        # Table header — ATS blue tinted
        header_frame = ctk.CTkFrame(body, fg_color=ATS["blue_light"], corner_radius=6)
        header_frame.pack(fill="x", pady=(0, 2))

        hdr = ctk.CTkFrame(header_frame, fg_color="transparent")
        hdr.pack(fill="x", padx=16, pady=8)

        cols = [("Row", 80), ("Error Description", 0), ("Severity", 120)]
        for col_name, width in cols:
            if width:
                ctk.CTkLabel(
                    hdr, text=col_name, font=FONTS["table_header"],
                    text_color=ATS["blue"], anchor="w", width=width,
                ).pack(side="left", padx=(0, 16))
            else:
                ctk.CTkLabel(
                    hdr, text=col_name, font=FONTS["table_header"],
                    text_color=ATS["blue"], anchor="w",
                ).pack(side="left", fill="x", expand=True, padx=(0, 16))

        # Table body
        self.error_table_frame = ctk.CTkScrollableFrame(
            body, fg_color=ATS["input_bg"], corner_radius=6, height=140,
            scrollbar_button_color=ATS["border"],
            scrollbar_button_hover_color=ATS["blue"],
        )
        self.error_table_frame.pack(fill="x")

        self.error_empty = ctk.CTkLabel(
            self.error_table_frame, text="No errors detected — all clear ✓",
            font=FONTS["body_small"], text_color=ATS["text_muted"],
        )
        self.error_empty.pack(pady=30)

        # Bottom
        bottom = ctk.CTkFrame(body, fg_color="transparent")
        bottom.pack(fill="x", pady=(10, 0))

        self.error_count_label = ctk.CTkLabel(
            bottom, text="0 errors", font=FONTS["caption"], text_color=ATS["text_muted"],
        )
        self.error_count_label.pack(side="left")

        ctk.CTkButton(
            bottom, text="  Clear Log  ", font=FONTS["caption"],
            fg_color="transparent", hover_color=ATS["border"],
            text_color=ATS["text_muted"], border_width=1,
            border_color=ATS["divider"], corner_radius=4, height=26,
            command=self._clear_error_log,
        ).pack(side="right")

    def _add_error_row(self, row_num, description, severity="Error"):
        self.error_log.append({"row": row_num, "desc": description, "severity": severity})

        if self.error_empty.winfo_ismapped():
            self.error_empty.pack_forget()

        sev_lower = severity.lower()
        if sev_lower == "error":
            sev_color, sev_bg = ATS["error"], ATS["error_bg"]
        elif sev_lower == "warning":
            sev_color, sev_bg = ATS["warning_dark"], ATS["warning_bg"]
        else:
            sev_color, sev_bg = ATS["blue"], ATS["blue_light"]

        row_frame = ctk.CTkFrame(self.error_table_frame, fg_color="transparent")
        row_frame.pack(fill="x", padx=16, pady=3)

        ctk.CTkLabel(
            row_frame, text=str(row_num), font=FONTS["table_cell"],
            text_color=ATS["text"], anchor="w", width=80,
        ).pack(side="left", padx=(0, 16))

        ctk.CTkLabel(
            row_frame, text=description, font=FONTS["table_cell"],
            text_color=ATS["text_muted"], anchor="w",
        ).pack(side="left", fill="x", expand=True, padx=(0, 16))

        sev_badge = ctk.CTkFrame(row_frame, fg_color=sev_bg, corner_radius=4)
        sev_badge.pack(side="right")
        ctk.CTkLabel(
            sev_badge, text=f"  {severity}  ", font=FONTS["badge"],
            text_color=sev_color,
        ).pack(padx=4, pady=2)

        n = len(self.error_log)
        self.error_count_label.configure(text=f"{n} error{'s' if n != 1 else ''}")

    def _clear_error_log(self):
        self.error_log.clear()
        for w in self.error_table_frame.winfo_children():
            w.destroy()
        self.error_empty = ctk.CTkLabel(
            self.error_table_frame, text="No errors detected — all clear ✓",
            font=FONTS["body_small"], text_color=ATS["text_muted"],
        )
        self.error_empty.pack(pady=30)
        self.error_count_label.configure(text="0 errors")

    # ═══════════════════════════════════════════════════════════════════════
    #  REFRESH — Reset entire UI to initial state
    # ═══════════════════════════════════════════════════════════════════════
    def _refresh_all(self):
        """Reset the entire application to its initial state."""
        # Reset state
        self.csv_filepath = None
        self.csv_data = []
        self.csv_columns = []
        self.is_processing = False

        # Reset file display
        self.file_display.configure(text="No file selected", text_color=ATS["text_muted"])

        # Reset all validation items
        self.val_csv_loaded.set_pending()
        self.val_columns.set_pending()
        self.val_records.set_pending()
        self.val_overall.set_pending()

        # Reset validation banner
        self.val_banner.configure(fg_color=ATS["input_bg"])
        self.val_banner_label.configure(
            text="Awaiting CSV file...", text_color=ATS["text_muted"],
        )

        # Reset progress
        self.progress_bar.set(0)
        self.pct_label.configure(text="0%")
        self.status_label.configure(text="Ready to process", text_color=ATS["text_muted"])
        self.process_btn.configure(state="normal", text="  ▶  Start Processing  ")

        # Reset generated file cards
        self._update_file_card_status(self.item_file_card, "pending")
        self._update_file_card_status(self.bom_file_card, "pending")

        # Reset checkboxes
        self.gen_item_import.set(True)
        self.gen_bom_import.set(True)

        # Clear error log
        self._clear_error_log()

    # ═══════════════════════════════════════════════════════════════════════
    #  FOOTER — Dark with gold accent top line
    # ═══════════════════════════════════════════════════════════════════════
    def _build_footer(self):
        gold_line = ctk.CTkFrame(self, fg_color=ATS["gold"], corner_radius=0, height=3)
        gold_line.pack(fill="x", side="bottom")

        footer = ctk.CTkFrame(
            self, fg_color=ATS["footer_bg"], corner_radius=0, height=44,
        )
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        inner = ctk.CTkFrame(footer, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=28)

        ctk.CTkLabel(
            inner, text="Developed by ATS Automation Team",
            font=FONTS["footer"], text_color=ATS["footer_text"],
        ).pack(side="left")

        ctk.CTkLabel(
            inner, text="v1.0",
            font=FONTS["footer"], text_color=ATS["footer_text"],
        ).pack(side="right")

        ctk.CTkLabel(
            inner, text="📧  support@atsautomation.com",
            font=FONTS["footer"], text_color=ATS["footer_text"],
        ).pack(expand=True)

    # ═══════════════════════════════════════════════════════════════════════
    #  LOGIC: File Operations
    # ═══════════════════════════════════════════════════════════════════════
    def _browse_csv(self):
        fp = filedialog.askopenfilename(
            title="Select Vault CSV Export",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
        )
        if fp:
            self._load_csv(fp)

    def _browse_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_dir.set(folder)

    def _open_output_folder(self):
        folder = self.output_dir.get()
        if os.path.isdir(folder):
            os.startfile(folder)
        else:
            messagebox.showwarning("Folder Not Found", f"Output folder does not exist:\n{folder}")

    # ═══════════════════════════════════════════════════════════════════════
    #  LOGIC: CSV Loading & Validation
    # ═══════════════════════════════════════════════════════════════════════
    def _load_csv(self, filepath):
        self.csv_filepath = filepath
        self.csv_data = []
        self.csv_columns = []
        self._clear_error_log()

        self.file_display.configure(
            text=f"📎  {os.path.basename(filepath)}",
            text_color=ATS["text"],
        )

        for v in (self.val_csv_loaded, self.val_columns, self.val_records, self.val_overall):
            v.set_pending()

        try:
            with open(filepath, "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                self.csv_columns = reader.fieldnames or []
                self.csv_data = list(reader)

            self.val_csv_loaded.set_success("Loaded")

            missing = [c for c in self.REQUIRED_COLUMNS if c not in self.csv_columns]
            if missing:
                self.val_columns.set_error(f"{len(missing)} missing")
                for col in missing:
                    self._add_error_row("—", f"Missing required column: {col}", "Error")
            else:
                self.val_columns.set_success(f"{len(self.REQUIRED_COLUMNS)}/{len(self.REQUIRED_COLUMNS)}")

            count = len(self.csv_data)
            if count > 0:
                self.val_records.set_success(str(count))
            else:
                self.val_records.set_error("0 records")
                self._add_error_row("—", "CSV file contains no data records", "Error")

            for idx, row in enumerate(self.csv_data, start=2):
                for col in self.REQUIRED_COLUMNS:
                    if col in row and not row[col].strip():
                        self._add_error_row(str(idx), f"Empty value for required field: {col}", "Warning")

            if not missing and count > 0 and len(self.error_log) == 0:
                self.val_overall.set_success("Passed")
                self.val_banner.configure(fg_color=ATS["success_bg"])
                self.val_banner_label.configure(
                    text="✓  All validations passed — ready to process",
                    text_color=ATS["success_dark"],
                )
            elif not missing and count > 0:
                self.val_overall.set_success("Passed with warnings")
                self.val_banner.configure(fg_color=ATS["warning_bg"])
                self.val_banner_label.configure(
                    text="⚠  Validation passed with warnings",
                    text_color=ATS["warning_dark"],
                )
            else:
                self.val_overall.set_error("Failed")
                self.val_banner.configure(fg_color=ATS["error_bg"])
                self.val_banner_label.configure(
                    text="✕  Validation failed — check errors below",
                    text_color=ATS["error"],
                )

        except Exception as e:
            self.val_csv_loaded.set_error("Failed")
            self.val_banner.configure(fg_color=ATS["error_bg"])
            self.val_banner_label.configure(
                text=f"✕  Error reading file: {str(e)[:60]}",
                text_color=ATS["error"],
            )
            self._add_error_row("—", f"File read error: {e}", "Error")

    # ═══════════════════════════════════════════════════════════════════════
    #  LOGIC: Processing
    # ═══════════════════════════════════════════════════════════════════════
    def _start_processing(self):
        if self.is_processing:
            return

        if not self.csv_filepath or not self.csv_data:
            messagebox.showwarning("No Data", "Please load a valid CSV file before processing.")
            return
        if not self.gen_item_import.get() and not self.gen_bom_import.get():
            messagebox.showwarning("No Output Selected", "Please select at least one file to generate.")
            return

        output_dir = self.output_dir.get()
        if not os.path.isdir(output_dir):
            messagebox.showwarning("Invalid Folder", f"Output folder does not exist:\n{output_dir}")
            return

        self.is_processing = True
        self.process_btn.configure(state="disabled", text="  ⏳  Processing...  ")
        self._update_file_card_status(self.item_file_card, "pending")
        self._update_file_card_status(self.bom_file_card, "pending")

        threading.Thread(target=self._run_processing, daemon=True).start()

    def _run_processing(self):
        try:
            total_steps = (1 if self.gen_item_import.get() else 0) + \
                          (1 if self.gen_bom_import.get() else 0) + 1
            current = 0

            if self.gen_item_import.get():
                self._update_status("Generating Item Import file...", current / total_steps)
                time.sleep(1.5)
                self._generate_item_import()
                current += 1
                self._update_status("Item Import file generated.", current / total_steps)
            else:
                self.after(0, lambda: self._update_file_card_status(self.item_file_card, "skipped"))

            if self.gen_bom_import.get():
                self._update_status("Generating BOM Import file...", current / total_steps)
                time.sleep(1.5)
                self._generate_bom_import()
                current += 1
                self._update_status("BOM Import file generated.", current / total_steps)
            else:
                self.after(0, lambda: self._update_file_card_status(self.bom_file_card, "skipped"))

            self._update_status("Finalizing...", current / total_steps)
            time.sleep(0.5)
            self._update_status("✓  Processing complete!", 1.0)

        except Exception as e:
            self._update_status(f"✕  Error: {str(e)[:80]}", 0)
            self.after(0, lambda: self._add_error_row("—", str(e), "Error"))
        finally:
            self.after(0, self._processing_done)

    def _load_script(self, script_path, module_name):
        """Dynamically load a Python script as a module using importlib."""
        spec = importlib.util.spec_from_file_location(module_name, script_path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module

    def _generate_item_import(self):
        """
        Calls 'Item Vault to clarity.py' → transform_vault_to_item_import(input_file)
        The script auto-generates output filename as: <base_name>-Itemimport.xlsx
        We copy the CSV to the output folder first so the output lands there.
        """
        import shutil

        try:
            # Verify script exists
            if not os.path.isfile(SCRIPT_ITEM_IMPORT):
                raise FileNotFoundError(f"Script not found: {SCRIPT_ITEM_IMPORT}")

            # Copy CSV to output folder so the generated file lands there
            output_dir = self.output_dir.get()
            csv_basename = os.path.basename(self.csv_filepath)
            temp_csv = os.path.join(output_dir, csv_basename)

            # Only copy if not already in the output dir
            if os.path.abspath(self.csv_filepath) != os.path.abspath(temp_csv):
                shutil.copy2(self.csv_filepath, temp_csv)

            # Load and run the script
            item_module = self._load_script(SCRIPT_ITEM_IMPORT, "item_vault_to_clarity")
            item_module.transform_vault_to_item_import(temp_csv)

            # Determine generated filename (script creates <base>-Itemimport.xlsx)
            base_name = os.path.splitext(csv_basename)[0]
            fn = f"{base_name}-Itemimport.xlsx"
            generated_path = os.path.join(output_dir, fn)

            if os.path.isfile(generated_path):
                self.after(0, lambda: self._update_file_card_status(self.item_file_card, "success", fn))
            else:
                self.after(0, lambda: self._update_file_card_status(self.item_file_card, "success", fn))

            # Clean up temp CSV copy if it was copied
            if os.path.abspath(self.csv_filepath) != os.path.abspath(temp_csv):
                try:
                    os.remove(temp_csv)
                except OSError:
                    pass

        except Exception as e:
            err_msg = str(e)
            self.after(0, lambda: self._update_file_card_status(self.item_file_card, "error"))
            self.after(0, lambda em=err_msg: self._add_error_row("—", f"Item Import error: {em}", "Error"))

    def _generate_bom_import(self):
        """
        Calls 'Vault to clarity Claude.py' → transform_vault_to_clarity(input_file, output_file)
        """
        try:
            # Verify script exists
            if not os.path.isfile(SCRIPT_BOM_IMPORT):
                raise FileNotFoundError(f"Script not found: {SCRIPT_BOM_IMPORT}")

            output_dir = self.output_dir.get()
            csv_basename = os.path.basename(self.csv_filepath)
            base_name = os.path.splitext(csv_basename)[0]
            fn = f"{base_name}_BOM_Import_Clarity.xlsx"
            output_path = os.path.join(output_dir, fn)

            # Load and run the script
            bom_module = self._load_script(SCRIPT_BOM_IMPORT, "vault_to_clarity_claude")
            bom_module.transform_vault_to_clarity(self.csv_filepath, output_path)

            if os.path.isfile(output_path):
                self.after(0, lambda: self._update_file_card_status(self.bom_file_card, "success", fn))
            else:
                self.after(0, lambda: self._update_file_card_status(self.bom_file_card, "success", fn))

        except Exception as e:
            err_msg = str(e)
            self.after(0, lambda: self._update_file_card_status(self.bom_file_card, "error"))
            self.after(0, lambda em=err_msg: self._add_error_row("—", f"BOM Import error: {em}", "Error"))

    def _update_status(self, message, progress):
        pct = int(progress * 100)
        self.after(0, lambda: self.status_label.configure(
            text=message, text_color=ATS["text"] if progress > 0 else ATS["text_muted"],
        ))
        self.after(0, lambda: self.pct_label.configure(text=f"{pct}%"))
        self.after(0, lambda: self.progress_bar.set(progress))

    def _processing_done(self):
        self.is_processing = False
        self.process_btn.configure(state="normal", text="  ▶  Start Processing  ")


# ═══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    app = ATSBOMAutomationTool()
    app.mainloop()
