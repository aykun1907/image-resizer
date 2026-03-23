import sys
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image
import json
import ctypes
import time
import logging
from datetime import datetime
try:
    import pillow_avif
except ImportError:
    pass  # Pillow 10+ has built-in AVIF support
from PIL import ImageOps


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def format_file_size(size_bytes):
    for unit in ('B', 'KB', 'MB', 'GB'):
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"


def get_unique_path(output_path):
    """If output_path already exists, append (1), (2), etc."""
    if not os.path.exists(output_path):
        return output_path
    directory = os.path.dirname(output_path)
    name, ext = os.path.splitext(os.path.basename(output_path))
    counter = 1
    while True:
        candidate = os.path.join(directory, f"{name} ({counter}){ext}")
        if not os.path.exists(candidate):
            return candidate
        counter += 1


class ImageResizerApp:
    def __init__(self, root, initial_files=None):
        logging.info("Initializing ImageResizerApp")
        self.root = root
        self.root.title("Image Resizer")
        self.root.resizable(False, False)
        self.image_paths = list(initial_files) if initial_files else []
        self.output_directory = None

        # Load config
        self.config_path = os.path.join(os.getenv('LOCALAPPDATA'), 'ImageResizer', 'config.json')
        os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
        self.load_config()

        # Style
        self.style = ttk.Style()
        self.style.configure('Header.TLabel', font=('Segoe UI', 14))
        self.style.configure('Info.TLabel', font=('Segoe UI', 9), foreground='#666666')
        self.style.configure('Status.TLabel', font=('Segoe UI', 9))

        # Scrollable canvas for the whole window
        self.canvas = tk.Canvas(root)
        self.scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.build_ui()

        # Size the window after building UI
        self.root.update_idletasks()
        frame_height = self.scrollable_frame.winfo_reqheight()
        screen_height = self.root.winfo_screenheight()
        max_height = min(frame_height + 40, screen_height - 100)
        width = 440
        self.root.geometry(f"{width}x{max_height}")

        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (max_height // 2)
        self.root.geometry(f"+{x}+{y}")

        # Keyboard shortcuts
        self.root.bind('<Escape>', lambda e: self.root.quit())
        self.root.bind('<Return>', lambda e: self.process())

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def build_ui(self):
        main = ttk.Frame(self.scrollable_frame, padding="20 15 20 15")
        main.pack(fill=tk.BOTH, expand=True)

        # =====================================================
        # Context Menu Status (compact bar at top)
        # =====================================================
        ctx_frame = ttk.Frame(main)
        ctx_frame.pack(fill=tk.X, pady=(0, 12))

        self.status_label = ttk.Label(ctx_frame, text="", style='Status.TLabel')
        self.status_label.pack(side=tk.LEFT)

        self.toggle_btn = ttk.Button(ctx_frame, text="", width=18, command=self.toggle_integration)
        self.toggle_btn.pack(side=tk.RIGHT)

        self.update_status()

        ttk.Separator(main, orient='horizontal').pack(fill=tk.X, pady=(0, 12))

        # =====================================================
        # Files Section
        # =====================================================
        files_frame = ttk.LabelFrame(main, text="Images", padding="10 10 10 10")
        files_frame.pack(fill=tk.X, pady=(0, 10))

        self.files_info_label = ttk.Label(files_frame, text="", wraplength=370)
        self.files_info_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        btn_container = ttk.Frame(files_frame)
        btn_container.pack(side=tk.RIGHT)

        self.clear_btn = ttk.Button(btn_container, text="Clear", width=6, command=self.clear_files)
        self.clear_btn.pack(side=tk.RIGHT, padx=(5, 0))

        ttk.Button(btn_container, text="Browse...", width=9, command=self.browse_files).pack(side=tk.RIGHT)

        self.update_files_display()

        # =====================================================
        # Size Options
        # =====================================================
        size_frame = ttk.LabelFrame(main, text="Size Options", padding="10 10 10 10")
        size_frame.pack(fill=tk.X, pady=(0, 10))

        self.size_mode = tk.StringVar(value="percentage")
        ttk.Radiobutton(size_frame, text="Keep original size (convert only)",
                       variable=self.size_mode, value="original",
                       command=self.update_size_fields).pack(anchor=tk.W)
        ttk.Radiobutton(size_frame, text="Scale by percentage",
                       variable=self.size_mode, value="percentage",
                       command=self.update_size_fields).pack(anchor=tk.W)
        ttk.Radiobutton(size_frame, text="Fit to max width",
                       variable=self.size_mode, value="fit_width",
                       command=self.update_size_fields).pack(anchor=tk.W)
        ttk.Radiobutton(size_frame, text="Fit to max height",
                       variable=self.size_mode, value="fit_height",
                       command=self.update_size_fields).pack(anchor=tk.W)
        ttk.Radiobutton(size_frame, text="Exact dimensions",
                       variable=self.size_mode, value="dimensions",
                       command=self.update_size_fields).pack(anchor=tk.W)

        self.size_container = ttk.Frame(size_frame)
        self.size_container.pack(fill=tk.X, pady=(10, 0))

        # Percentage
        self.scale_frame = ttk.Frame(self.size_container)
        ttk.Label(self.scale_frame, text="Scale:").pack(side=tk.LEFT)
        self.scale_var = tk.StringVar(value="50")
        ttk.Entry(self.scale_frame, textvariable=self.scale_var, width=6).pack(side=tk.LEFT, padx=5)
        ttk.Label(self.scale_frame, text="%").pack(side=tk.LEFT)

        # Width only
        self.width_only_frame = ttk.Frame(self.size_container)
        ttk.Label(self.width_only_frame, text="Max width:").pack(side=tk.LEFT)
        self.fit_width_var = tk.StringVar(value="1280")
        ttk.Entry(self.width_only_frame, textvariable=self.fit_width_var, width=6).pack(side=tk.LEFT, padx=5)
        ttk.Label(self.width_only_frame, text="px  (height auto)", style='Info.TLabel').pack(side=tk.LEFT)

        # Height only
        self.height_only_frame = ttk.Frame(self.size_container)
        ttk.Label(self.height_only_frame, text="Max height:").pack(side=tk.LEFT)
        self.fit_height_var = tk.StringVar(value="720")
        ttk.Entry(self.height_only_frame, textvariable=self.fit_height_var, width=6).pack(side=tk.LEFT, padx=5)
        ttk.Label(self.height_only_frame, text="px  (width auto)", style='Info.TLabel').pack(side=tk.LEFT)

        # Both dimensions
        self.dim_frame = ttk.Frame(self.size_container)
        ttk.Label(self.dim_frame, text="Width:").pack(side=tk.LEFT)
        self.width_var = tk.StringVar(value="800")
        ttk.Entry(self.dim_frame, textvariable=self.width_var, width=6).pack(side=tk.LEFT, padx=5)
        ttk.Label(self.dim_frame, text="Height:").pack(side=tk.LEFT, padx=(10, 0))
        self.height_var = tk.StringVar(value="600")
        ttk.Entry(self.dim_frame, textvariable=self.height_var, width=6).pack(side=tk.LEFT, padx=5)

        # Aspect ratio (only for exact dimensions)
        self.aspect_frame = ttk.Frame(size_frame)
        self.aspect_mode = tk.StringVar(value="maintain")
        ttk.Radiobutton(self.aspect_frame, text="Maintain aspect ratio",
                       variable=self.aspect_mode, value="maintain").pack(anchor=tk.W)
        ttk.Radiobutton(self.aspect_frame, text="Stretch to fit",
                       variable=self.aspect_mode, value="stretch").pack(anchor=tk.W)
        ttk.Radiobutton(self.aspect_frame, text="Crop to fit",
                       variable=self.aspect_mode, value="crop").pack(anchor=tk.W)

        # =====================================================
        # Output Location
        # =====================================================
        folder_frame = ttk.LabelFrame(main, text="Output Location", padding="10 10 10 10")
        folder_frame.pack(fill=tk.X, pady=(0, 10))

        self.output_mode = tk.StringVar(value="same")
        ttk.Radiobutton(folder_frame, text="Same folder as original",
                       variable=self.output_mode, value="same",
                       command=self.update_folder_display).pack(anchor=tk.W)
        ttk.Radiobutton(folder_frame, text="Custom folder",
                       variable=self.output_mode, value="custom",
                       command=self.update_folder_display).pack(anchor=tk.W)

        self.folder_container = ttk.Frame(folder_frame)
        self.folder_container.pack(fill=tk.X, pady=(5, 0))
        self.folder_path_var = tk.StringVar(value="")
        self.folder_path_label = ttk.Label(self.folder_container, textvariable=self.folder_path_var,
                                         wraplength=300)
        self.folder_button = ttk.Button(self.folder_container, text="Browse...",
                                      command=self.browse_folder)

        # =====================================================
        # Output Options
        # =====================================================
        quality_frame = ttk.LabelFrame(main, text="Output Options", padding="10 10 10 10")
        quality_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(quality_frame, text="Format:").pack(anchor=tk.W)
        self.format_var = tk.StringVar(value="same")
        formats = [("Keep original", "same"),
                  ("JPEG", "JPEG"),
                  ("PNG", "PNG"),
                  ("WebP", "WebP"),
                  ("AVIF", "AVIF")]

        format_frame = ttk.Frame(quality_frame)
        format_frame.pack(fill=tk.X, pady=(5, 10))
        for text, value in formats:
            ttk.Radiobutton(format_frame, text=text,
                           variable=self.format_var, value=value).pack(anchor=tk.W)

        # Quality slider
        self.quality_container = ttk.Frame(quality_frame)
        self.quality_container.pack(fill=tk.X)

        quality_label_frame = ttk.Frame(self.quality_container)
        quality_label_frame.pack(fill=tk.X)
        ttk.Label(quality_label_frame, text="Quality:").pack(side=tk.LEFT)
        self.quality_percentage = ttk.Label(quality_label_frame, text="85%")
        self.quality_percentage.pack(side=tk.RIGHT)

        self.quality_var = tk.IntVar(value=85)
        self.quality_slider = ttk.Scale(self.quality_container, from_=1, to=100,
                                 orient=tk.HORIZONTAL,
                                 variable=self.quality_var,
                                 command=self.update_quality_percentage)
        self.quality_slider.pack(fill=tk.X, pady=(5, 0))

        self.quality_note = ttk.Label(self.quality_container, text="", style='Info.TLabel')

        # --- JPEG Options ---
        self.jpeg_frame = ttk.LabelFrame(quality_frame, text="JPEG Options", padding="5 5 5 5")

        ttk.Label(self.jpeg_frame, text="Chroma subsampling:").pack(anchor=tk.W)
        self.jpeg_subsampling_var = tk.StringVar(value="4:2:0")
        sub_frame = ttk.Frame(self.jpeg_frame)
        sub_frame.pack(fill=tk.X, pady=(2, 5))
        ttk.Radiobutton(sub_frame, text="4:2:0 (smaller file, default)",
                       variable=self.jpeg_subsampling_var, value="4:2:0").pack(anchor=tk.W)
        ttk.Radiobutton(sub_frame, text="4:4:4 (sharper colors, larger file)",
                       variable=self.jpeg_subsampling_var, value="4:4:4").pack(anchor=tk.W)

        self.jpeg_optimize_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.jpeg_frame, text="Optimize Huffman coding (smaller, no quality loss)",
                       variable=self.jpeg_optimize_var).pack(anchor=tk.W)

        # --- WebP Options ---
        self.webp_frame = ttk.LabelFrame(quality_frame, text="WebP Options", padding="5 5 5 5")

        method_label_frame = ttk.Frame(self.webp_frame)
        method_label_frame.pack(fill=tk.X)
        ttk.Label(method_label_frame, text="Compression effort:").pack(side=tk.LEFT)
        self.webp_method_label = ttk.Label(method_label_frame, text="4 (default)")
        self.webp_method_label.pack(side=tk.RIGHT)

        self.webp_method_var = tk.IntVar(value=4)
        ttk.Scale(self.webp_frame, from_=0, to=6,
                 orient=tk.HORIZONTAL,
                 variable=self.webp_method_var,
                 command=self.update_webp_method_label).pack(fill=tk.X, pady=(5, 0))
        ttk.Label(self.webp_frame, text="Higher = slower but better compression at same quality",
                 style='Info.TLabel').pack(anchor=tk.W, pady=(3, 0))

        # --- AVIF Options ---
        self.avif_frame = ttk.LabelFrame(quality_frame, text="AVIF Options", padding="5 5 5 5")

        speed_label_frame = ttk.Frame(self.avif_frame)
        speed_label_frame.pack(fill=tk.X)
        ttk.Label(speed_label_frame, text="Speed:").pack(side=tk.LEFT)
        self.speed_percentage = ttk.Label(speed_label_frame, text="6 (balanced)")
        self.speed_percentage.pack(side=tk.RIGHT)

        self.speed_var = tk.IntVar(value=6)
        ttk.Scale(self.avif_frame, from_=0, to=10,
                 orient=tk.HORIZONTAL,
                 variable=self.speed_var,
                 command=self.update_speed_label).pack(fill=tk.X, pady=(5, 0))

        self.format_var.trace_add('write', self.update_format_options)

        # =====================================================
        # Output Preview
        # =====================================================
        preview_frame = ttk.LabelFrame(main, text="Output Preview", padding="10 10 10 10")
        preview_frame.pack(fill=tk.X, pady=(0, 10))

        self.preview_label = ttk.Label(preview_frame, wraplength=380)
        self.preview_label.pack(fill=tk.X)

        self.size_mode.trace_add('write', self.update_preview)
        self.scale_var.trace_add('write', self.update_preview)
        self.fit_width_var.trace_add('write', self.update_preview)
        self.fit_height_var.trace_add('write', self.update_preview)
        self.width_var.trace_add('write', self.update_preview)
        self.height_var.trace_add('write', self.update_preview)
        self.aspect_mode.trace_add('write', self.update_preview)
        self.format_var.trace_add('write', self.update_preview)

        # =====================================================
        # Bottom: checkboxes + action button
        # =====================================================
        bottom_frame = ttk.Frame(main)
        bottom_frame.pack(fill=tk.X, pady=(5, 0))

        self.save_default = tk.BooleanVar(value=False)
        ttk.Checkbutton(bottom_frame, text="Save as default settings",
                       variable=self.save_default).pack(anchor=tk.W)

        self.keep_open_var = tk.BooleanVar(value=self.config.get('keep_open', False))
        ttk.Checkbutton(bottom_frame, text="Keep open after processing",
                       variable=self.keep_open_var).pack(anchor=tk.W, pady=(0, 10))

        btn_frame = ttk.Frame(bottom_frame)
        btn_frame.pack(fill=tk.X)

        self.process_btn = ttk.Button(btn_frame, text="Resize", command=self.process)
        self.process_btn.pack(side=tk.RIGHT)

        # =====================================================
        # Apply defaults from config
        # =====================================================
        self.size_mode.set(self.config['default_size_mode'])
        self.scale_var.set(str(self.config['default_scale']))
        self.width_var.set(str(self.config['default_width']))
        self.height_var.set(str(self.config['default_height']))
        self.aspect_mode.set(self.config['default_aspect_mode'])
        self.quality_var.set(self.config['default_quality'])
        self.format_var.set(self.config['default_format'])

        self.update_size_fields()
        self.update_folder_display()
        self.update_format_options()

    # ---------------------------------------------------------
    # Files management
    # ---------------------------------------------------------
    def browse_files(self):
        files = filedialog.askopenfilenames(
            title="Select images",
            filetypes=[
                ("Images", "*.jpg *.jpeg *.png *.gif *.bmp *.webp *.avif"),
                ("All files", "*.*")
            ]
        )
        if files:
            self.image_paths = list(files)
            self.update_files_display()

    def clear_files(self):
        self.image_paths = []
        self.update_files_display()

    def update_files_display(self):
        count = len(self.image_paths)
        if count == 0:
            self.files_info_label.config(text="No images selected")
            self.clear_btn.state(['disabled'])
        else:
            try:
                total_size = sum(os.path.getsize(p) for p in self.image_paths)
                if count == 1:
                    with Image.open(self.image_paths[0]) as img:
                        name = os.path.basename(self.image_paths[0])
                        self.files_info_label.config(
                            text=f"{name}  |  {img.width}x{img.height}  |  {format_file_size(total_size)}")
                else:
                    self.files_info_label.config(
                        text=f"{count} files  |  {format_file_size(total_size)} total")
            except Exception:
                self.files_info_label.config(text=f"{count} file{'s' if count > 1 else ''} selected")
            self.clear_btn.state(['!disabled'])

    # ---------------------------------------------------------
    # Size fields
    # ---------------------------------------------------------
    def update_size_fields(self):
        mode = self.size_mode.get()
        # Hide all input frames first
        self.scale_frame.pack_forget()
        self.width_only_frame.pack_forget()
        self.height_only_frame.pack_forget()
        self.dim_frame.pack_forget()
        self.aspect_frame.pack_forget()

        if mode == "original":
            self.process_btn.config(text="Convert")
            self.root.title("Image Converter")
        elif mode == "percentage":
            self.scale_frame.pack(fill=tk.X)
            self.process_btn.config(text="Resize")
            self.root.title("Image Resizer")
        elif mode == "fit_width":
            self.width_only_frame.pack(fill=tk.X)
            self.process_btn.config(text="Resize")
            self.root.title("Image Resizer")
        elif mode == "fit_height":
            self.height_only_frame.pack(fill=tk.X)
            self.process_btn.config(text="Resize")
            self.root.title("Image Resizer")
        else:  # dimensions
            self.dim_frame.pack(fill=tk.X)
            self.aspect_frame.pack(fill=tk.X, pady=(10, 0))
            self.process_btn.config(text="Resize")
            self.root.title("Image Resizer")

    # ---------------------------------------------------------
    # Output folder
    # ---------------------------------------------------------
    def update_folder_display(self):
        if self.output_mode.get() == "custom":
            self.folder_button.pack(side=tk.RIGHT, padx=(5, 0))
            self.folder_path_label.pack(side=tk.LEFT, fill=tk.X)
        else:
            self.folder_button.pack_forget()
            self.folder_path_label.pack_forget()
            self.folder_path_var.set("")
            self.output_directory = None

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_directory = folder
            self.folder_path_var.set(folder)

    # ---------------------------------------------------------
    # Format-specific UI updates
    # ---------------------------------------------------------
    def update_format_options(self, *args):
        fmt = self.format_var.get()
        self.jpeg_frame.pack_forget()
        self.webp_frame.pack_forget()
        self.avif_frame.pack_forget()

        if fmt == "JPEG":
            self.jpeg_frame.pack(fill=tk.X, pady=(10, 0))
        elif fmt == "WebP":
            self.webp_frame.pack(fill=tk.X, pady=(10, 0))
        elif fmt == "AVIF":
            self.avif_frame.pack(fill=tk.X, pady=(10, 0))

        if fmt == "PNG":
            self.quality_slider.state(['disabled'])
            self.quality_percentage.config(text="N/A")
            self.quality_note.config(text="PNG is lossless (quality slider has no effect)")
            self.quality_note.pack(pady=(3, 0))
        else:
            self.quality_slider.state(['!disabled'])
            self.quality_percentage.config(text=f"{self.quality_var.get()}%")
            self.quality_note.pack_forget()

    def update_quality_percentage(self, *args):
        self.quality_percentage.config(text=f"{self.quality_var.get()}%")

    def update_speed_label(self, *args):
        speed = self.speed_var.get()
        if speed <= 2:
            desc = "slower, better quality"
        elif speed <= 4:
            desc = "slow"
        elif speed <= 6:
            desc = "balanced"
        elif speed <= 8:
            desc = "fast"
        else:
            desc = "faster, lower quality"
        self.speed_percentage.config(text=f"{speed} ({desc})")

    def update_webp_method_label(self, *args):
        method = self.webp_method_var.get()
        if method <= 1:
            desc = "fastest"
        elif method <= 2:
            desc = "fast"
        elif method <= 3:
            desc = "moderate"
        elif method <= 4:
            desc = "default"
        elif method <= 5:
            desc = "slow, better compression"
        else:
            desc = "slowest, best compression"
        self.webp_method_label.config(text=f"{method} ({desc})")

    def update_preview(self, *args):
        try:
            mode = self.size_mode.get()
            if mode == 'original':
                suffix = "_converted"
            elif mode == 'percentage':
                scale = float(self.scale_var.get())
                suffix = f"_{int(scale)}p"
            elif mode == 'fit_width':
                w = self.fit_width_var.get()
                suffix = f"_{w}w"
            elif mode == 'fit_height':
                h = self.fit_height_var.get()
                suffix = f"_{h}h"
            else:
                width = self.width_var.get()
                height = self.height_var.get()
                aspect = self.aspect_mode.get()
                if aspect == 'crop':
                    suffix = f"_{width}x{height}_crop"
                elif aspect == 'stretch':
                    suffix = f"_{width}x{height}_stretch"
                else:
                    suffix = f"_{width}x{height}"

            format_str = self.format_var.get()
            ext = '.[original]' if format_str == 'same' else f'.{format_str.lower()}'
            preview = f"example{suffix}{ext}"
            self.preview_label.config(text=f"Output filename pattern: {preview}")
        except (ValueError, TypeError):
            self.preview_label.config(text="Invalid input values")

    # ---------------------------------------------------------
    # Validation
    # ---------------------------------------------------------
    def validate_inputs(self):
        if not self.image_paths:
            messagebox.showerror("Error", "No images selected. Use Browse to pick files.")
            return False

        mode = self.size_mode.get()
        if mode == "original":
            pass
        elif mode == "percentage":
            try:
                scale = float(self.scale_var.get())
                if not 1 <= scale <= 1000:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Error", "Please enter a valid scale percentage (1-1000)")
                return False
        elif mode == "fit_width":
            try:
                w = int(self.fit_width_var.get())
                if w <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Error", "Please enter a valid width (positive number)")
                return False
        elif mode == "fit_height":
            try:
                h = int(self.fit_height_var.get())
                if h <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Error", "Please enter a valid height (positive number)")
                return False
        else:
            try:
                width = int(self.width_var.get())
                height = int(self.height_var.get())
                if width <= 0 or height <= 0:
                    raise ValueError
                if width > 10000 or height > 10000:
                    if not messagebox.askyesno("Warning",
                        "The dimensions are very large. This might consume a lot of memory. Continue?"):
                        return False
            except ValueError:
                messagebox.showerror("Error", "Please enter valid dimensions (positive numbers)")
                return False

        if self.output_mode.get() == "custom" and not self.output_directory:
            messagebox.showerror("Error", "Please select an output folder")
            return False

        return True

    # ---------------------------------------------------------
    # Process images
    # ---------------------------------------------------------
    def process(self):
        if not self.validate_inputs():
            return

        mode = self.size_mode.get()
        options = {
            'size_mode': mode,
            'scale': float(self.scale_var.get()) if mode == "percentage" else None,
            'fit_width': int(self.fit_width_var.get()) if mode == "fit_width" else None,
            'fit_height': int(self.fit_height_var.get()) if mode == "fit_height" else None,
            'width': int(self.width_var.get()) if mode == "dimensions" else None,
            'height': int(self.height_var.get()) if mode == "dimensions" else None,
            'aspect_mode': self.aspect_mode.get(),
            'format': self.format_var.get(),
            'quality': self.quality_var.get(),
            'avif_speed': self.speed_var.get(),
            'webp_method': self.webp_method_var.get(),
            'jpeg_subsampling': self.jpeg_subsampling_var.get(),
            'jpeg_optimize': self.jpeg_optimize_var.get(),
            'output_dir': self.output_directory if self.output_mode.get() == "custom" else None,
        }

        if self.save_default.get():
            self.save_as_default(options)

        # Persist keep_open preference
        keep_open = self.keep_open_var.get()
        if keep_open != self.config.get('keep_open', False):
            self.config['keep_open'] = keep_open
            self.save_config()

        try:
            total = len(self.image_paths)
            progress = None
            if total > 1:
                progress = self._create_progress(total)

            for idx, img_path in enumerate(self.image_paths, 1):
                try:
                    logging.info(f"Processing image: {img_path}")
                    if progress:
                        progress['label'].config(text=f"Processing {os.path.basename(img_path)} ({idx}/{total})...")
                        progress['bar']['value'] = idx
                        progress['win'].update()

                    with Image.open(img_path) as img:
                        img = ImageOps.exif_transpose(img)

                        size_mode = options['size_mode']
                        if size_mode == 'original':
                            resized = img.copy()
                            size_suffix = "_converted"
                        elif size_mode == 'percentage':
                            scale = options['scale'] / 100
                            new_width = int(img.width * scale)
                            new_height = int(img.height * scale)
                            size_suffix = f"_{int(scale * 100)}p"
                            resized = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                        elif size_mode == 'fit_width':
                            target_w = options['fit_width']
                            ratio = target_w / img.width
                            new_width = target_w
                            new_height = int(img.height * ratio)
                            size_suffix = f"_{new_width}w"
                            resized = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                        elif size_mode == 'fit_height':
                            target_h = options['fit_height']
                            ratio = target_h / img.height
                            new_width = int(img.width * ratio)
                            new_height = target_h
                            size_suffix = f"_{new_height}h"
                            resized = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                        else:  # dimensions
                            new_width = options['width']
                            new_height = options['height']

                            if options['aspect_mode'] == 'maintain':
                                width_ratio = new_width / img.width
                                height_ratio = new_height / img.height
                                ratio = min(width_ratio, height_ratio)
                                new_width = int(img.width * ratio)
                                new_height = int(img.height * ratio)
                                size_suffix = f"_{new_width}x{new_height}"
                                resized = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                            elif options['aspect_mode'] == 'crop':
                                width_ratio = new_width / img.width
                                height_ratio = new_height / img.height
                                ratio = max(width_ratio, height_ratio)
                                resize_width = int(img.width * ratio)
                                resize_height = int(img.height * ratio)
                                resized = img.resize((resize_width, resize_height), Image.Resampling.LANCZOS)
                                left = (resize_width - new_width) // 2
                                top = (resize_height - new_height) // 2
                                resized = resized.crop((left, top, left + new_width, top + new_height))
                                size_suffix = f"_{new_width}x{new_height}_crop"
                            else:
                                resized = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                                size_suffix = f"_{new_width}x{new_height}_stretch"

                        output_format = options['format']
                        if output_format == 'same':
                            output_format = img.format
                            # Fallback: if Pillow can't detect format, guess from extension
                            if not output_format:
                                ext_map = {
                                    '.jpg': 'JPEG', '.jpeg': 'JPEG',
                                    '.png': 'PNG', '.gif': 'GIF',
                                    '.bmp': 'BMP', '.webp': 'WebP',
                                    '.avif': 'AVIF',
                                }
                                src_ext = os.path.splitext(img_path)[1].lower()
                                output_format = ext_map.get(src_ext, 'PNG')

                        if options['output_dir']:
                            basename = os.path.basename(img_path)
                            name, _ = os.path.splitext(basename)
                            ext = '.' + output_format.lower()
                            output_path = os.path.join(options['output_dir'], f"{name}{size_suffix}{ext}")
                        else:
                            name, _ = os.path.splitext(img_path)
                            ext = '.' + output_format.lower()
                            output_path = f"{name}{size_suffix}{ext}"

                        # Don't overwrite the source file
                        if os.path.normpath(os.path.abspath(output_path)) == \
                           os.path.normpath(os.path.abspath(img_path)):
                            name_part, ext_part = os.path.splitext(output_path)
                            output_path = f"{name_part}_new{ext_part}"

                        output_path = get_unique_path(output_path)

                        # Build save options
                        save_kwargs = {}
                        if output_format == 'JPEG':
                            save_kwargs['quality'] = options['quality']
                            save_kwargs['optimize'] = options['jpeg_optimize']
                            save_kwargs['subsampling'] = 0 if options['jpeg_subsampling'] == '4:4:4' else 2
                        elif output_format == 'WebP':
                            save_kwargs['quality'] = options['quality']
                            save_kwargs['method'] = options['webp_method']
                        elif output_format == 'PNG':
                            save_kwargs['optimize'] = True
                        elif output_format == 'AVIF':
                            quality = options['quality']
                            if quality >= 93:
                                save_kwargs.update({
                                    'quality': quality,
                                    'speed': min(4, options['avif_speed']),
                                    'chroma': 1.0, 'lossless': False
                                })
                            elif quality >= 81:
                                save_kwargs.update({
                                    'quality': quality,
                                    'speed': options['avif_speed'],
                                    'chroma': 1.0, 'lossless': False
                                })
                            else:
                                save_kwargs.update({
                                    'quality': quality,
                                    'speed': options['avif_speed'],
                                    'chroma': 0.75, 'lossless': False
                                })

                        if output_format == 'JPEG' and resized.mode in ('RGBA', 'P', 'LA'):
                            resized = resized.convert('RGB')
                        elif output_format == 'AVIF' and resized.mode not in ("RGB", "RGBA"):
                            resized = resized.convert("RGBA")

                        resized.save(output_path, output_format, **save_kwargs)

                except Exception as e:
                    logging.exception(f"Failed to process {img_path}")
                    messagebox.showerror("Error", f"Failed to process {img_path}: {str(e)}")

            if progress:
                progress['win'].destroy()

            # Open output folder
            output_dir = options['output_dir'] or os.path.dirname(self.image_paths[0])
            logging.info(f"Opening output directory: {output_dir}")
            os.startfile(output_dir)

        except Exception as e:
            logging.exception("Error during image processing")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

        # After processing
        if keep_open:
            self.image_paths = []
            self.update_files_display()
        else:
            self.root.quit()

    def _create_progress(self, total):
        win = tk.Toplevel(self.root)
        win.title("Processing...")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        frame = ttk.Frame(win, padding="20 15 20 15")
        frame.pack(fill=tk.BOTH, expand=True)

        label = ttk.Label(frame, text=f"Processing image 1 of {total}...")
        label.pack(pady=(0, 10))

        bar = ttk.Progressbar(frame, length=300, mode='determinate', maximum=total)
        bar.pack()

        win.update_idletasks()
        w = win.winfo_reqwidth() + 40
        h = win.winfo_reqheight() + 20
        x = (win.winfo_screenwidth() // 2) - (w // 2)
        y = (win.winfo_screenheight() // 2) - (h // 2)
        win.geometry(f"{w}x{h}+{x}+{y}")

        return {'win': win, 'label': label, 'bar': bar}

    # ---------------------------------------------------------
    # Context menu integration
    # ---------------------------------------------------------
    def check_integration_status(self):
        try:
            sendto_path = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'SendTo')
            shortcut_path = os.path.join(sendto_path, 'Resize Images.lnk')
            return os.path.exists(shortcut_path)
        except Exception as e:
            logging.error(f"Error checking integration status: {e}")
            return False

    def update_status(self):
        is_active = self.check_integration_status()
        if is_active:
            self.status_label.config(text="Send To menu: Active")
            self.toggle_btn.config(text="Disable Integration")
        else:
            self.status_label.config(text="Send To menu: Inactive")
            self.toggle_btn.config(text="Enable Integration")

    def toggle_integration(self):
        if not is_admin():
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas",
                sys.executable if getattr(sys, 'frozen', False) else sys.argv[0],
                f'"{sys.argv[0]}" {"--disable" if self.check_integration_status() else "--enable"}',
                None, 1
            )
            self.root.after(1000, self.update_status)
            return

        if self.check_integration_status():
            self.remove_context_menu()
        else:
            self.install_context_menu()
        self.update_status()

    def install_context_menu(self):
        exe_path = sys.executable if getattr(sys, 'frozen', False) else sys.argv[0]
        exe_path = os.path.abspath(exe_path)
        try:
            sendto_path = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'SendTo')
            shortcut_path = os.path.join(sendto_path, 'Resize Images.lnk')

            ps_script = f'''
            $WshShell = New-Object -comObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut("{shortcut_path}")
            $Shortcut.TargetPath = "{exe_path}"
            $Shortcut.Arguments = "resize"
            $Shortcut.Save()
            '''
            import subprocess
            subprocess.run(['powershell', '-Command', ps_script], capture_output=True, text=True)

            logging.info(f"Created SendTo shortcut at {shortcut_path}")
            messagebox.showinfo("Success",
                "Image Resizer has been added to the SendTo menu!\n"
                "Select images > right-click > Send To > Resize Images")
        except Exception as e:
            logging.exception("Failed to create SendTo shortcut")
            messagebox.showerror("Error", f"Failed to add SendTo menu: {str(e)}")

    def remove_context_menu(self):
        try:
            sendto_path = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'SendTo')
            shortcut_path = os.path.join(sendto_path, 'Resize Images.lnk')
            if os.path.exists(shortcut_path):
                os.remove(shortcut_path)
                logging.info(f"Removed SendTo shortcut: {shortcut_path}")
            messagebox.showinfo("Success", "Image Resizer has been removed from the SendTo menu!")
        except Exception as e:
            logging.exception("Failed to remove SendTo shortcut")
            messagebox.showerror("Error", f"Failed to remove SendTo menu: {str(e)}")

    # ---------------------------------------------------------
    # Config
    # ---------------------------------------------------------
    def load_config(self):
        try:
            with open(self.config_path, 'r') as f:
                self.config = json.load(f)
        except FileNotFoundError:
            self.config = {}
            self.save_config()

        defaults = {
            'default_width': 800,
            'default_height': 600,
            'default_scale': 50,
            'default_quality': 85,
            'default_format': 'same',
            'default_size_mode': 'percentage',
            'default_aspect_mode': 'maintain',
            'keep_open': False,
        }
        for key, value in defaults.items():
            if key not in self.config:
                self.config[key] = value

    def save_config(self):
        with open(self.config_path, 'w') as f:
            json.dump(self.config, f, indent=4)

    def save_as_default(self, options):
        self.config.update({
            'default_width': options['width'] or 800,
            'default_height': options['height'] or 600,
            'default_scale': options['scale'] or 50,
            'default_quality': options['quality'],
            'default_format': options['format'],
            'default_size_mode': options['size_mode'],
            'default_aspect_mode': options['aspect_mode'],
        })
        self.save_config()


# =============================================================
# Logging & entry point
# =============================================================

def setup_logging():
    log_dir = os.path.join(os.getenv('LOCALAPPDATA'), 'ImageResizer', 'logs')
    os.makedirs(log_dir, exist_ok=True)

    try:
        log_files = sorted(
            [os.path.join(log_dir, f) for f in os.listdir(log_dir) if f.endswith('.log')],
            key=os.path.getmtime
        )
        for old_log in log_files[:-20]:
            os.remove(old_log)
    except OSError:
        pass

    log_file = os.path.join(log_dir, f'imageresizer_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    logging.info(f"Logging started. Log file: {log_file}")
    logging.info(f"Python version: {sys.version}")
    logging.info(f"Command line arguments: {sys.argv}")


class SingleInstance:
    def __init__(self):
        self.lockfile = os.path.join(os.getenv('TEMP'), 'imageresizer.lock')
        self.locked = False

    def __enter__(self):
        try:
            if os.path.exists(self.lockfile):
                if time.time() - os.path.getmtime(self.lockfile) > 30:
                    os.remove(self.lockfile)
            with open(self.lockfile, 'x') as f:
                f.write(str(os.getpid()))
            self.locked = True
            return True
        except FileExistsError:
            return False

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.locked and os.path.exists(self.lockfile):
            try:
                os.remove(self.lockfile)
            except OSError:
                pass


def main():
    setup_logging()
    logging.info("Application starting")

    if len(sys.argv) > 1:
        logging.info(f"Running with command: {sys.argv[1]}")

        if sys.argv[1] == "resize":
            with SingleInstance() as single:
                if single:
                    # Collect image files from args
                    image_files = []
                    if len(sys.argv) > 2:
                        for arg in sys.argv[2:]:
                            path = arg.strip('"').strip("'")
                            path = os.path.normpath(path)
                            if os.path.isfile(path) and any(path.lower().endswith(ext) for ext in
                                  ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.avif')):
                                image_files.append(path)

                    logging.info(f"Found {len(image_files)} valid image files")

                    root = tk.Tk()
                    app = ImageResizerApp(root, initial_files=image_files)
                    root.mainloop()
                else:
                    logging.warning("Another instance is already running")

        elif sys.argv[1] in ["--enable", "--disable"]:
            root = tk.Tk()
            root.withdraw()
            app = ImageResizerApp(root)
            if sys.argv[1] == "--enable":
                app.install_context_menu()
            else:
                app.remove_context_menu()
            root.destroy()
    else:
        # Standalone — open with no files, user can browse
        root = tk.Tk()
        app = ImageResizerApp(root)
        root.mainloop()


if __name__ == "__main__":
    main()
