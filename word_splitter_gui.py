"""
word_splitter_gui.py
====================
Modern Graphical User Interface (GUI) for the Word Document Splitter.
Built with Tkinter and styled with a sleek modern theme.
"""

import os
import sys
import threading
import subprocess
from typing import Optional
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from word_splitter_core import WordSplitter


# Color Palette (Tailwind Slate / Indigo Theme)
THEME = {
    "bg_main": "#0f172a",       # Slate 900
    "bg_card": "#1e293b",       # Slate 800
    "bg_card_hover": "#243248",
    "bg_input": "#0f172a",      # Slate 900
    "border": "#334155",        # Slate 700
    "border_focus": "#6366f1",  # Indigo 500
    "text_main": "#f8fafc",     # Slate 50
    "text_muted": "#94a3b8",    # Slate 400
    "text_highlight": "#38bdf8",# Sky 400
    "btn_primary": "#4f46e5",   # Indigo 600
    "btn_primary_hover": "#4338ca",
    "btn_success": "#059669",   # Emerald 600
    "btn_success_hover": "#047857",
    "btn_secondary": "#334155", # Slate 700
    "btn_secondary_hover": "#475569",
    "chip_bg": "#334155",
    "chip_active": "#4f46e5",
    "accent_blue": "#2563eb",
    "badge_bg": "#1e1b4b",      # Deep Indigo
    "badge_text": "#a5b4fc",
}

FONT_FAMILY = "Segoe UI" if sys.platform == "win32" else "Helvetica"


def format_size(bytes_num: int) -> str:
    """Format bytes to human readable string."""
    for unit in ["B", "KB", "MB", "GB"]:
        if bytes_num < 1024.0:
            return f"{bytes_num:.1f} {unit}" if unit != "B" else f"{bytes_num} B"
        bytes_num /= 1024.0
    return f"{bytes_num:.1f} TB"


def open_path_in_explorer(path: str):
    """Open a file or directory in the system file manager."""
    if not os.path.exists(path):
        return
    if sys.platform == "win32":
        if os.path.isdir(path):
            os.startfile(path)
        else:
            subprocess.run(f'explorer /select,"{os.path.abspath(path)}"', shell=True)
    elif sys.platform == "darwin":
        subprocess.run(["open", path])
    else:
        subprocess.run(["xdg-open", path])


class ModernCard(tk.Frame):
    """Card container with subtle border and padding."""
    def __init__(self, parent, **kwargs):
        super().__init__(
            parent,
            bg=THEME["bg_card"],
            highlightbackground=THEME["border"],
            highlightthickness=1,
            padx=16,
            pady=16,
            **kwargs
        )


class WordSplitterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Word File Splitter")
        self.geometry("780x760")
        self.minsize(680, 640)
        self.configure(bg=THEME["bg_main"])

        # State
        self.current_file: Optional[str] = None
        self.doc_analysis: Optional[dict] = None
        self.is_processing = False
        self.parts_var = tk.IntVar(value=3)
        self.method_var = tk.StringVar(value="Equal Memory / File Size (Recommended)")
        self.open_folder_var = tk.BooleanVar(value=True)

        self.setup_styles()
        self.build_ui()

    def setup_styles(self):
        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        # Configure ttk Progressbar
        self.style.configure(
            "Horizontal.TProgressbar",
            troughcolor=THEME["bg_input"],
            background=THEME["btn_primary"],
            bordercolor=THEME["border"],
            lightcolor=THEME["btn_primary"],
            darkcolor=THEME["btn_primary"],
            thickness=12
        )

        # Configure Combobox
        self.style.configure(
            "TCombobox",
            fieldbackground=THEME["bg_input"],
            background=THEME["bg_card"],
            foreground=THEME["text_main"],
            arrowcolor=THEME["text_muted"],
            bordercolor=THEME["border"],
            darkcolor=THEME["bg_input"],
            lightcolor=THEME["bg_input"]
        )
        self.style.map(
            "TCombobox",
            fieldbackground=[("readonly", THEME["bg_input"])],
            selectbackground=[("readonly", THEME["btn_primary"])],
            selectforeground=[("readonly", THEME["text_main"])]
        )

    def build_ui(self):
        # Main scrollable canvas container to ensure it works on any screen resolution
        canvas = tk.Canvas(self, bg=THEME["bg_main"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scroll_frame = tk.Frame(canvas, bg=THEME["bg_main"])

        self.scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas_window = canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")

        def on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)

        canvas.bind("<Configure>", on_canvas_configure)
        canvas.configure(xscrollcommand=None, yscrollcommand=scrollbar.set)

        # Mousewheel scroll support
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", on_mousewheel)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Content container inside scroll_frame
        container = tk.Frame(self.scroll_frame, bg=THEME["bg_main"], padx=24, pady=20)
        container.pack(fill="both", expand=True)

        # 1. Header Banner
        self.create_header(container)

        # 2. File Selection Card
        self.create_file_card(container)

        # 3. Document Summary Card (Initially hidden until document loaded)
        self.create_summary_card(container)

        # 4. Split Settings Card
        self.create_settings_card(container)

        # 5. Output Destination Card
        self.create_output_card(container)

        # 6. Action Button & Progress
        self.create_action_section(container)

        # 7. Results Section
        self.create_results_section(container)

    def create_header(self, parent):
        header_frame = tk.Frame(parent, bg=THEME["bg_main"], pady=4)
        header_frame.pack(fill="x", pady=(0, 16))

        title_label = tk.Label(
            header_frame,
            text="📄 Word Document Splitter",
            font=(FONT_FAMILY, 20, "bold"),
            fg=THEME["text_main"],
            bg=THEME["bg_main"]
        )
        title_label.pack(anchor="w")

        subtitle_label = tk.Label(
            header_frame,
            text="Split any .docx document into N balanced parts with 100% formatting, table, and style retention.",
            font=(FONT_FAMILY, 10),
            fg=THEME["text_muted"],
            bg=THEME["bg_main"]
        )
        subtitle_label.pack(anchor="w", pady=(2, 0))

    def create_file_card(self, parent):
        card = ModernCard(parent)
        card.pack(fill="x", pady=(0, 14))

        title = tk.Label(
            card,
            text="1. Select Word Document (.docx)",
            font=(FONT_FAMILY, 11, "bold"),
            fg=THEME["text_main"],
            bg=THEME["bg_card"]
        )
        title.pack(anchor="w", pady=(0, 10))

        input_row = tk.Frame(card, bg=THEME["bg_card"])
        input_row.pack(fill="x")

        self.file_entry = tk.Entry(
            input_row,
            font=(FONT_FAMILY, 10),
            bg=THEME["bg_input"],
            fg=THEME["text_main"],
            insertbackground=THEME["text_main"],
            highlightbackground=THEME["border"],
            highlightcolor=THEME["border_focus"],
            highlightthickness=1,
            relief="flat",
            bd=6
        )
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.file_entry.bind("<Return>", lambda e: self.on_path_entered())

        browse_btn = tk.Button(
            input_row,
            text="Browse File...",
            font=(FONT_FAMILY, 10, "bold"),
            bg=THEME["btn_primary"],
            fg=THEME["text_main"],
            activebackground=THEME["btn_primary_hover"],
            activeforeground=THEME["text_main"],
            relief="flat",
            bd=0,
            padx=16,
            pady=8,
            cursor="hand2",
            command=self.browse_file
        )
        browse_btn.pack(side="right")

    def create_summary_card(self, parent):
        self.summary_card = ModernCard(parent)
        # Initially not packed until file is selected

        title = tk.Label(
            self.summary_card,
            text="📊 Document Overview",
            font=(FONT_FAMILY, 11, "bold"),
            fg=THEME["text_main"],
            bg=THEME["bg_card"]
        )
        title.pack(anchor="w", pady=(0, 10))

        self.stats_grid = tk.Frame(self.summary_card, bg=THEME["bg_card"])
        self.stats_grid.pack(fill="x")

        # Stats badges
        self.stat_widgets = {}
        labels = [
            ("words", "Total Words"),
            ("blocks", "Content Blocks"),
            ("paragraphs", "Paragraphs"),
            ("tables", "Tables"),
            ("size", "File Size"),
            ("sections", "Sections"),
        ]

        for i, (key, label_text) in enumerate(labels):
            box = tk.Frame(
                self.stats_grid,
                bg=THEME["bg_input"],
                highlightbackground=THEME["border"],
                highlightthickness=1,
                padx=10,
                pady=8
            )
            box.grid(row=i // 3, column=i % 3, sticky="nsew", padx=4, pady=4)
            self.stats_grid.grid_columnconfigure(i % 3, weight=1)

            val_label = tk.Label(
                box,
                text="--",
                font=(FONT_FAMILY, 12, "bold"),
                fg=THEME["text_highlight"],
                bg=THEME["bg_input"]
            )
            val_label.pack(anchor="w")

            title_lbl = tk.Label(
                box,
                text=label_text,
                font=(FONT_FAMILY, 8),
                fg=THEME["text_muted"],
                bg=THEME["bg_input"]
            )
            title_lbl.pack(anchor="w")
            self.stat_widgets[key] = val_label

    def create_settings_card(self, parent):
        card = ModernCard(parent)
        card.pack(fill="x", pady=(0, 14))

        title = tk.Label(
            card,
            text="2. Split Configuration",
            font=(FONT_FAMILY, 11, "bold"),
            fg=THEME["text_main"],
            bg=THEME["bg_card"]
        )
        title.pack(anchor="w", pady=(0, 12))

        # Row 1: Number of Segments
        parts_row = tk.Frame(card, bg=THEME["bg_card"])
        parts_row.pack(fill="x", pady=(0, 10))

        parts_label = tk.Label(
            parts_row,
            text="Number of Segments:",
            font=(FONT_FAMILY, 10),
            fg=THEME["text_main"],
            bg=THEME["bg_card"]
        )
        parts_label.pack(side="left", padx=(0, 12))

        # Minus button
        minus_btn = tk.Button(
            parts_row,
            text="−",
            font=(FONT_FAMILY, 12, "bold"),
            bg=THEME["btn_secondary"],
            fg=THEME["text_main"],
            activebackground=THEME["btn_secondary_hover"],
            activeforeground=THEME["text_main"],
            relief="flat",
            bd=0,
            width=3,
            cursor="hand2",
            command=self.decrement_parts
        )
        minus_btn.pack(side="left")

        # Parts Entry
        self.parts_entry = tk.Entry(
            parts_row,
            textvariable=self.parts_var,
            font=(FONT_FAMILY, 11, "bold"),
            justify="center",
            width=5,
            bg=THEME["bg_input"],
            fg=THEME["text_highlight"],
            insertbackground=THEME["text_main"],
            highlightbackground=THEME["border"],
            highlightthickness=1,
            relief="flat",
            bd=4
        )
        self.parts_entry.pack(side="left", padx=4)
        self.parts_var.trace_add("write", lambda *args: self.update_preview())

        # Plus button
        plus_btn = tk.Button(
            parts_row,
            text="+",
            font=(FONT_FAMILY, 12, "bold"),
            bg=THEME["btn_secondary"],
            fg=THEME["text_main"],
            activebackground=THEME["btn_secondary_hover"],
            activeforeground=THEME["text_main"],
            relief="flat",
            bd=0,
            width=3,
            cursor="hand2",
            command=self.increment_parts
        )
        plus_btn.pack(side="left", padx=(0, 16))

        # Quick preset chips
        presets_frame = tk.Frame(parts_row, bg=THEME["bg_card"])
        presets_frame.pack(side="left")

        for num in [2, 3, 4, 5, 10]:
            btn = tk.Button(
                presets_frame,
                text=f"{num} Segments",
                font=(FONT_FAMILY, 8, "bold"),
                bg=THEME["chip_bg"],
                fg=THEME["text_main"],
                activebackground=THEME["chip_active"],
                activeforeground=THEME["text_main"],
                relief="flat",
                bd=0,
                padx=8,
                pady=4,
                cursor="hand2",
                command=lambda n=num: self.set_parts(n)
            )
            btn.pack(side="left", padx=2)

        # Row 2: Split Method
        method_row = tk.Frame(card, bg=THEME["bg_card"])
        method_row.pack(fill="x", pady=(8, 4))

        method_label = tk.Label(
            method_row,
            text="Splitting Method:",
            font=(FONT_FAMILY, 10),
            fg=THEME["text_main"],
            bg=THEME["bg_card"]
        )
        method_label.pack(side="left", padx=(0, 12))

        self.method_combo = ttk.Combobox(
            method_row,
            textvariable=self.method_var,
            state="readonly",
            font=(FONT_FAMILY, 9),
            width=38,
            values=[
                "Equal Memory / File Size (Recommended)",
                "Balanced Word Count",
                "Equal Content Blocks (Paragraphs & Tables)",
                "By Chapters / Headings",
            ]
        )
        self.method_combo.pack(side="left")
        self.method_combo.bind("<<ComboboxSelected>>", lambda e: self.update_preview())

        # Estimated breakdown preview
        self.preview_label = tk.Label(
            card,
            text="Estimated per part: --",
            font=(FONT_FAMILY, 9),
            fg=THEME["text_highlight"],
            bg=THEME["bg_card"]
        )
        self.preview_label.pack(anchor="w", pady=(8, 0))

    def create_output_card(self, parent):
        card = ModernCard(parent)
        card.pack(fill="x", pady=(0, 14))

        title = tk.Label(
            card,
            text="3. Output Destination",
            font=(FONT_FAMILY, 11, "bold"),
            fg=THEME["text_main"],
            bg=THEME["bg_card"]
        )
        title.pack(anchor="w", pady=(0, 10))

        out_row = tk.Frame(card, bg=THEME["bg_card"])
        out_row.pack(fill="x")

        self.output_entry = tk.Entry(
            out_row,
            font=(FONT_FAMILY, 10),
            bg=THEME["bg_input"],
            fg=THEME["text_main"],
            insertbackground=THEME["text_main"],
            highlightbackground=THEME["border"],
            highlightcolor=THEME["border_focus"],
            highlightthickness=1,
            relief="flat",
            bd=6
        )
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        change_out_btn = tk.Button(
            out_row,
            text="Change Folder...",
            font=(FONT_FAMILY, 9, "bold"),
            bg=THEME["btn_secondary"],
            fg=THEME["text_main"],
            activebackground=THEME["btn_secondary_hover"],
            activeforeground=THEME["text_main"],
            relief="flat",
            bd=0,
            padx=12,
            pady=6,
            cursor="hand2",
            command=self.browse_output_dir
        )
        change_out_btn.pack(side="right")

        # Checkbox: open folder on finish
        chk = tk.Checkbutton(
            card,
            text="Open folder automatically when finished",
            variable=self.open_folder_var,
            font=(FONT_FAMILY, 9),
            bg=THEME["bg_card"],
            fg=THEME["text_muted"],
            activebackground=THEME["bg_card"],
            activeforeground=THEME["text_main"],
            selectcolor=THEME["bg_input"],
            cursor="hand2"
        )
        chk.pack(anchor="w", pady=(8, 0))

    def create_action_section(self, parent):
        self.action_frame = tk.Frame(parent, bg=THEME["bg_main"])
        self.action_frame.pack(fill="x", pady=(4, 14))

        self.split_btn = tk.Button(
            self.action_frame,
            text="✂️  Split Document Now",
            font=(FONT_FAMILY, 12, "bold"),
            bg=THEME["btn_primary"],
            fg=THEME["text_main"],
            activebackground=THEME["btn_primary_hover"],
            activeforeground=THEME["text_main"],
            relief="flat",
            bd=0,
            pady=12,
            cursor="hand2",
            command=self.start_split_thread
        )
        self.split_btn.pack(fill="x")

        # Progress bar
        self.progress_bar = ttk.Progressbar(
            self.action_frame,
            style="Horizontal.TProgressbar",
            mode="determinate"
        )
        self.status_label = tk.Label(
            self.action_frame,
            text="Ready to split",
            font=(FONT_FAMILY, 9),
            fg=THEME["text_muted"],
            bg=THEME["bg_main"]
        )

    def create_results_section(self, parent):
        self.results_card = ModernCard(parent)
        # Hidden initially until completion

        title_row = tk.Frame(self.results_card, bg=THEME["bg_card"])
        title_row.pack(fill="x", pady=(0, 10))

        title = tk.Label(
            title_row,
            text="🎉 Split Completed Successfully!",
            font=(FONT_FAMILY, 12, "bold"),
            fg=THEME["text_highlight"],
            bg=THEME["bg_card"]
        )
        title.pack(side="left")

        self.open_dir_btn = tk.Button(
            title_row,
            text="📂 Open Output Folder",
            font=(FONT_FAMILY, 9, "bold"),
            bg=THEME["btn_success"],
            fg=THEME["text_main"],
            activebackground=THEME["btn_success_hover"],
            activeforeground=THEME["text_main"],
            relief="flat",
            bd=0,
            padx=12,
            pady=6,
            cursor="hand2",
            command=lambda: open_path_in_explorer(self.last_output_dir) if hasattr(self, "last_output_dir") else None
        )
        self.open_dir_btn.pack(side="right")

        self.parts_list_frame = tk.Frame(self.results_card, bg=THEME["bg_card"])
        self.parts_list_frame.pack(fill="x")

    # Helper & Event Handlers
    def increment_parts(self):
        try:
            val = self.parts_var.get()
            self.parts_var.set(val + 1)
        except Exception:
            self.parts_var.set(2)

    def decrement_parts(self):
        try:
            val = self.parts_var.get()
            if val > 2:
                self.parts_var.set(val - 1)
        except Exception:
            self.parts_var.set(2)

    def set_parts(self, n: int):
        self.parts_var.set(n)

    def browse_file(self):
        file_selected = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")]
        )
        if file_selected:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_selected)
            self.load_document(file_selected)

    def on_path_entered(self):
        path = self.file_entry.get().strip().strip('"').strip("'")
        if path and os.path.isfile(path):
            self.load_document(path)

    def browse_output_dir(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, folder)

    def load_document(self, path: str):
        if not os.path.isfile(path):
            messagebox.showerror("File Error", f"File not found:\n{path}")
            return
        if not path.lower().endswith(".docx"):
            messagebox.showerror("Unsupported File", "Please select a Microsoft Word (.docx) document.")
            return

        self.current_file = path
        try:
            self.doc_analysis = WordSplitter.analyze_document(path)
        except Exception as e:
            messagebox.showerror("Read Error", f"Could not read Word document:\n{e}")
            return

        # Update stats
        self.stat_widgets["words"].config(text=f"{self.doc_analysis['total_words']:,}")
        self.stat_widgets["blocks"].config(text=f"{self.doc_analysis['total_blocks']:,}")
        self.stat_widgets["paragraphs"].config(text=f"{self.doc_analysis['total_paragraphs']:,}")
        self.stat_widgets["tables"].config(text=f"{self.doc_analysis['total_tables']:,}")
        self.stat_widgets["size"].config(text=format_size(self.doc_analysis["file_size"]))
        self.stat_widgets["sections"].config(text=f"{self.doc_analysis['sections_count']}")

        # Display summary card
        self.summary_card.pack(fill="x", pady=(0, 14), after=self.summary_card.master.winfo_children()[1])

        # Suggest default output directory
        base_name, _ = os.path.splitext(self.doc_analysis["file_name"])
        default_dir = os.path.join(os.path.dirname(os.path.abspath(path)), f"{base_name}_split")
        self.output_entry.delete(0, tk.END)
        self.output_entry.insert(0, default_dir)

        # Update preview
        self.update_preview()

    def update_preview(self):
        if not self.doc_analysis:
            self.preview_label.config(text="Estimated per part: (select a file first)")
            return

        try:
            n_parts = self.parts_var.get()
        except Exception:
            return

        if n_parts <= 0:
            return

        total_words = self.doc_analysis["total_words"]
        total_blocks = self.doc_analysis["total_blocks"]

        actual_parts = min(n_parts, total_blocks)
        avg_words = total_words // actual_parts if actual_parts else 0
        avg_blocks = total_blocks // actual_parts if actual_parts else 0
        total_size = self.doc_analysis.get("file_size", 0)
        avg_size = total_size / actual_parts if actual_parts else 0

        method_text = self.method_var.get()
        if "Memory" in method_text or "File Size" in method_text:
            self.preview_label.config(
                text=f"Estimated per part: ~ {format_size(int(avg_size))} across {actual_parts} parts."
            )
        elif "Word Count" in method_text:
            self.preview_label.config(
                text=f"Estimated per part: ~ {avg_words:,} words across {actual_parts} parts."
            )
        elif "Blocks" in method_text:
            self.preview_label.config(
                text=f"Estimated per part: ~ {avg_blocks} content blocks across {actual_parts} parts."
            )
        else:
            headings_cnt = len(self.doc_analysis["headings"])
            self.preview_label.config(
                text=f"Split by headings: {headings_cnt} headings available."
            )

    def get_chosen_method_code(self) -> str:
        text = self.method_var.get()
        if "Memory" in text or "File Size" in text:
            return "memory"
        elif "Blocks" in text:
            return "blocks"
        elif "Headings" in text:
            return "headings"
        elif "Word" in text:
            return "words"
        return "memory"

    def start_split_thread(self):
        if self.is_processing:
            return

        # Check file
        file_path = self.file_entry.get().strip().strip('"').strip("'")
        if not file_path or not os.path.isfile(file_path):
            messagebox.showwarning("Missing File", "Please select a valid Word (.docx) file first.")
            return

        if not self.doc_analysis or self.current_file != file_path:
            self.load_document(file_path)
            if not self.doc_analysis:
                return

        # Check parts
        try:
            num_parts = self.parts_var.get()
            if num_parts < 1:
                raise ValueError
        except Exception:
            messagebox.showwarning("Invalid Input", "Please enter a valid number of parts (at least 1).")
            return

        # Check output directory
        out_dir = self.output_entry.get().strip().strip('"').strip("'")
        if not out_dir:
            base_name, _ = os.path.splitext(self.doc_analysis["file_name"])
            out_dir = os.path.join(os.path.dirname(os.path.abspath(file_path)), f"{base_name}_split")
            self.output_entry.insert(0, out_dir)

        # UI state during processing
        self.is_processing = True
        self.split_btn.config(state="disabled", text="⏳ Splitting Document...")
        self.progress_bar.pack(fill="x", pady=(10, 6))
        self.status_label.pack(anchor="w")
        self.status_label.config(text="Starting document processing...")
        self.progress_bar["value"] = 0

        # Hide old results if present
        self.results_card.pack_forget()

        # Run split in background thread
        thread = threading.Thread(
            target=self.run_split_worker,
            args=(file_path, num_parts, out_dir, self.get_chosen_method_code()),
            daemon=True
        )
        thread.start()

    def run_split_worker(self, file_path: str, num_parts: int, out_dir: str, method: str):
        def on_progress(current: int, total: int, part_path: str):
            pct = int((current / total) * 100)
            self.after(0, lambda: self.update_progress_ui(current, total, pct, part_path))

        try:
            result = WordSplitter.split_document(
                file_path=file_path,
                num_parts=num_parts,
                output_dir=out_dir,
                method=method,
                progress_callback=on_progress
            )
            self.after(0, lambda: self.on_split_success(result))
        except Exception as e:
            self.after(0, lambda: self.on_split_error(str(e)))

    def update_progress_ui(self, current: int, total: int, pct: int, part_path: str):
        self.progress_bar["value"] = pct
        fname = os.path.basename(part_path)
        self.status_label.config(text=f"Created Part {current} of {total}: {fname}")

    def on_split_success(self, result: dict):
        self.is_processing = False
        self.split_btn.config(state="normal", text="✂️  Split Document Now")
        self.progress_bar["value"] = 100
        self.status_label.config(text=f"✅ Finished! Generated {len(result['parts'])} parts successfully.")

        self.last_output_dir = result["output_directory"]

        # Populate Results Card
        for widget in self.parts_list_frame.winfo_children():
            widget.destroy()

        for p in result["parts"]:
            row = tk.Frame(
                self.parts_list_frame,
                bg=THEME["bg_input"],
                highlightbackground=THEME["border"],
                highlightthickness=1,
                padx=12,
                pady=8
            )
            row.pack(fill="x", pady=3)

            p_info = tk.Label(
                row,
                text=f"Part {p['part_number']}: {p['file_name']}",
                font=(FONT_FAMILY, 9, "bold"),
                fg=THEME["text_main"],
                bg=THEME["bg_input"]
            )
            p_info.pack(side="left")

            meta_info = tk.Label(
                row,
                text=f"{p['words']:,} words  •  {p['block_count']} blocks  •  {format_size(p['file_size'])}",
                font=(FONT_FAMILY, 8),
                fg=THEME["text_muted"],
                bg=THEME["bg_input"]
            )
            meta_info.pack(side="left", padx=12)

            open_doc_btn = tk.Button(
                row,
                text="Open File",
                font=(FONT_FAMILY, 8, "bold"),
                bg=THEME["btn_secondary"],
                fg=THEME["text_main"],
                activebackground=THEME["btn_secondary_hover"],
                activeforeground=THEME["text_main"],
                relief="flat",
                bd=0,
                padx=8,
                pady=2,
                cursor="hand2",
                command=lambda path=p["file_path"]: open_path_in_explorer(path)
            )
            open_doc_btn.pack(side="right")

        self.results_card.pack(fill="x", pady=(10, 20))

        if self.open_folder_var.get():
            open_path_in_explorer(result["output_directory"])

    def on_split_error(self, err_msg: str):
        self.is_processing = False
        self.split_btn.config(state="normal", text="✂️  Split Document Now")
        self.status_label.config(text=f"❌ Error during split: {err_msg}")
        messagebox.showerror("Splitting Error", f"An error occurred while splitting the document:\n{err_msg}")


def main():
    app = WordSplitterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
