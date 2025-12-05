"""
Bluetooth Label Printer App
A modern GUI for printing shipping labels via Bluetooth.
"""
import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import fitz  # PyMuPDF
from pathlib import Path
import threading
import logging
from io import BytesIO

from bt_printer import BluetoothPrinter, list_available_ports

__version__ = "1.0.0"
__author__ = "PrintToBTLabel Contributors"

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Set appearance
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class LabelPrinterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title(f"Bluetooth Label Printer v{__version__}")
        self.geometry("900x930")
        self.minsize(800, 600)
        
        # State
        self.pdf_path = None
        self.preview_image = None
        self.cropped_image = None
        self.printer = None
        self.selected_port = ctk.StringVar(value="COM5")
        
        # Label settings
        self.label_width = ctk.IntVar(value=101)  # 4 inches
        self.label_height = ctk.IntVar(value=152)  # 6 inches
        self.auto_crop = ctk.BooleanVar(value=True)
        self.auto_rotate = ctk.BooleanVar(value=True)
        self.flip_vertical = ctk.BooleanVar(value=False)
        self.invert_colors = ctk.BooleanVar(value=False)  # Off = invert (correct printing), On = don't invert
        self.manual_crop = ctk.BooleanVar(value=False)
        
        # Manual crop state
        self.crop_coords = None  # Applied crop (left, top, right, bottom) in original PDF coordinates
        self.pending_crop = None  # Selection before applying
        self.crop_start = None
        self.crop_rect = None
        self.original_pdf_image = None  # Full PDF image for manual crop selection
        self.crop_applied = False  # Whether crop has been applied
        
        self._create_ui()
        self._setup_crop_bindings()
        self._refresh_ports()
    
    def _create_ui(self):
        # Main container
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # Header
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        header.grid_columnconfigure(1, weight=1)
        
        title = ctk.CTkLabel(
            header, 
            text="📦 Label Printer",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title.grid(row=0, column=0, sticky="w")
        
        about_btn = ctk.CTkButton(
            header,
            text="ⓘ",
            width=30,
            height=30,
            command=self._show_about,
            fg_color="transparent",
            hover_color="#333333"
        )
        about_btn.grid(row=0, column=1, padx=10, sticky="w")
        
        # Connection frame
        conn_frame = ctk.CTkFrame(header)
        conn_frame.grid(row=0, column=2, sticky="e")
        
        self.port_dropdown = ctk.CTkComboBox(
            conn_frame,
            variable=self.selected_port,
            values=["COM5"],
            width=120
        )
        self.port_dropdown.grid(row=0, column=0, padx=5, pady=5)
        
        refresh_btn = ctk.CTkButton(
            conn_frame,
            text="↻",
            width=30,
            command=self._refresh_ports
        )
        refresh_btn.grid(row=0, column=1, padx=2, pady=5)
        
        self.status_label = ctk.CTkLabel(
            conn_frame,
            text="● Disconnected",
            text_color="#ff6b6b",
            font=ctk.CTkFont(size=12)
        )
        self.status_label.grid(row=0, column=2, padx=10, pady=5)
        
        # Main content area
        content = ctk.CTkFrame(self)
        content.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        content.grid_columnconfigure(0, weight=2)
        content.grid_columnconfigure(1, weight=1)
        content.grid_rowconfigure(0, weight=1)
        
        # Preview area (left)
        preview_frame = ctk.CTkFrame(content, fg_color="#1a1a2e")
        preview_frame.grid(row=0, column=0, padx=(10, 5), pady=10, sticky="nsew")
        preview_frame.grid_rowconfigure(1, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
        
        preview_header = ctk.CTkFrame(preview_frame, fg_color="transparent")
        preview_header.grid(row=0, column=0, padx=10, pady=5, sticky="ew")
        preview_header.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(
            preview_header,
            text="Preview",
            font=ctk.CTkFont(size=16, weight="bold")
        ).grid(row=0, column=0, sticky="w")
        
        self.filename_label = ctk.CTkLabel(
            preview_header,
            text="No file selected",
            text_color="gray"
        )
        self.filename_label.grid(row=0, column=1, padx=10, sticky="w")
        
        # Preview canvas with manual crop support
        # Use a regular frame with matching background instead of CTkFrame to avoid padding
        import tkinter as tk
        canvas_frame = tk.Frame(preview_frame, bg="#1a1a2e", highlightthickness=0)
        canvas_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        # Configure frame to center its contents
        canvas_frame.grid_rowconfigure(0, weight=1)
        canvas_frame.grid_columnconfigure(0, weight=1)
        
        # Use Canvas for manual crop interaction
        self.preview_canvas = ctk.CTkCanvas(
            canvas_frame,
            bg="#1a1a2e",
            highlightthickness=0,
            width=400,
            height=600
        )
        # Center canvas in frame (grid with weight centers it)
        self.preview_canvas.grid(row=0, column=0)
        
        # Placeholder text
        self.canvas_placeholder = ctk.CTkLabel(
            canvas_frame,
            text="Drop PDF here or click 'Select PDF'",
            fg_color="#1a1a2e",
            corner_radius=10,
            font=ctk.CTkFont(size=14)
        )
        self.canvas_placeholder.grid(row=0, column=0, sticky="nsew")
        
        # Bind mouse events for manual crop (will be set up after methods are defined)
        # We'll bind these in a separate method called after UI creation
        
        self.canvas_image_id = None
        self.original_image_size = None  # (width, height) of original before display scaling
        
        # Select button
        select_btn = ctk.CTkButton(
            preview_frame,
            text="📁 Select PDF",
            command=self._select_pdf,
            height=40,
            font=ctk.CTkFont(size=14)
        )
        select_btn.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        
        # Settings area (right)
        settings_frame = ctk.CTkFrame(content)
        settings_frame.grid(row=0, column=1, padx=(5, 10), pady=10, sticky="nsew")
        settings_frame.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(
            settings_frame,
            text="Settings",
            font=ctk.CTkFont(size=16, weight="bold")
        ).grid(row=0, column=0, padx=15, pady=(15, 10), sticky="w")
        
        # Label size
        size_frame = ctk.CTkFrame(settings_frame, fg_color="transparent")
        size_frame.grid(row=1, column=0, padx=15, pady=5, sticky="ew")
        size_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(size_frame, text="Label Size:").grid(row=0, column=0, sticky="w")
        
        size_options = ctk.CTkFrame(size_frame, fg_color="transparent")
        size_options.grid(row=1, column=0, columnspan=2, pady=5, sticky="ew")
        
        self.size_4x6_btn = ctk.CTkButton(
            size_options,
            text="4×6\"",
            width=70,
            command=lambda: self._set_size(101, 152),
            fg_color="#2d5a27" 
        )
        self.size_4x6_btn.grid(row=0, column=0, padx=2)
        
        self.size_4x4_btn = ctk.CTkButton(
            size_options,
            text="4×4\"",
            width=70,
            command=lambda: self._set_size(101, 101),
            fg_color="transparent",
            border_width=1
        )
        self.size_4x4_btn.grid(row=0, column=1, padx=2)
        
        self.size_2x1_btn = ctk.CTkButton(
            size_options,
            text="2×1\"",
            width=70,
            command=lambda: self._set_size(50, 25),
            fg_color="transparent",
            border_width=1
        )
        self.size_2x1_btn.grid(row=0, column=2, padx=2)
        
        self.size_custom_btn = ctk.CTkButton(
            size_options,
            text="Custom",
            width=70,
            command=self._set_custom_size,
            fg_color="transparent",
            border_width=1
        )
        self.size_custom_btn.grid(row=0, column=3, padx=2)
        
        # Custom size inputs
        custom_frame = ctk.CTkFrame(settings_frame, fg_color="transparent")
        custom_frame.grid(row=2, column=0, padx=15, pady=10, sticky="ew")
        
        ctk.CTkLabel(custom_frame, text="Width (mm):").grid(row=0, column=0, sticky="w")
        self.width_entry = ctk.CTkEntry(custom_frame, textvariable=self.label_width, width=80)
        self.width_entry.grid(row=0, column=1, padx=5)
        self.width_entry.bind("<FocusOut>", lambda e: self._on_custom_size_change())
        self.width_entry.bind("<Return>", lambda e: self._on_custom_size_change())
        
        ctk.CTkLabel(custom_frame, text="Height (mm):").grid(row=1, column=0, sticky="w", pady=(5, 0))
        self.height_entry = ctk.CTkEntry(custom_frame, textvariable=self.label_height, width=80)
        self.height_entry.grid(row=1, column=1, padx=5, pady=(5, 0))
        self.height_entry.bind("<FocusOut>", lambda e: self._on_custom_size_change())
        self.height_entry.bind("<Return>", lambda e: self._on_custom_size_change())
        
        # Options
        options_frame = ctk.CTkFrame(settings_frame, fg_color="transparent")
        options_frame.grid(row=3, column=0, padx=15, pady=10, sticky="ew")
        
        ctk.CTkLabel(
            options_frame,
            text="Options:",
            font=ctk.CTkFont(weight="bold")
        ).grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        self.crop_check = ctk.CTkCheckBox(
            options_frame,
            text="Auto-crop to label",
            variable=self.auto_crop,
            command=self._update_preview
        )
        self.crop_check.grid(row=1, column=0, sticky="w")
        
        self.rotate_check = ctk.CTkCheckBox(
            options_frame,
            text="Auto-rotate",
            variable=self.auto_rotate,
            command=self._update_preview
        )
        self.rotate_check.grid(row=2, column=0, sticky="w", pady=(5, 0))
        
        self.flip_check = ctk.CTkCheckBox(
            options_frame,
            text="Flip vertical",
            variable=self.flip_vertical,
            command=self._update_preview
        )
        self.flip_check.grid(row=3, column=0, sticky="w", pady=(5, 0))
        
        self.invert_check = ctk.CTkCheckBox(
            options_frame,
            text="Invert colors",
            variable=self.invert_colors,
            command=self._update_preview
        )
        self.invert_check.grid(row=4, column=0, sticky="w", pady=(5, 0))
        
        self.manual_crop_check = ctk.CTkCheckBox(
            options_frame,
            text="Manual crop",
            variable=self.manual_crop,
            command=self._toggle_manual_crop
        )
        self.manual_crop_check.grid(row=5, column=0, sticky="w", pady=(5, 0))
        
        # Crop buttons frame
        crop_btn_frame = ctk.CTkFrame(options_frame, fg_color="transparent")
        crop_btn_frame.grid(row=6, column=0, sticky="w", pady=(5, 0))
        
        self.apply_crop_btn = ctk.CTkButton(
            crop_btn_frame,
            text="Apply Crop",
            width=90,
            command=self._apply_crop,
            state="disabled",
            fg_color="#2563eb"
        )
        self.apply_crop_btn.grid(row=0, column=0, padx=(0, 5))
        
        self.clear_crop_btn = ctk.CTkButton(
            crop_btn_frame,
            text="Clear",
            width=60,
            command=self._clear_crop,
            state="disabled"
        )
        self.clear_crop_btn.grid(row=0, column=1)
        
        # Spacer
        settings_frame.grid_rowconfigure(4, weight=1)
        
        # Print button
        self.print_btn = ctk.CTkButton(
            settings_frame,
            text="🖨️ Print Label",
            command=self._print_label,
            height=50,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="#2d5a27",
            hover_color="#3d7a37"
        )
        self.print_btn.grid(row=5, column=0, padx=15, pady=15, sticky="sew")
        
        # Progress bar (hidden by default)
        self.progress = ctk.CTkProgressBar(settings_frame)
        self.progress.grid(row=6, column=0, padx=15, pady=(0, 15), sticky="ew")
        self.progress.set(0)
        self.progress.grid_remove()
        
        # Status bar
        status_bar = ctk.CTkFrame(self, height=30)
        status_bar.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
        
        self.status_text = ctk.CTkLabel(
            status_bar,
            text="Ready",
            font=ctk.CTkFont(size=12)
        )
        self.status_text.pack(side="left", padx=10, pady=5)
    
    def _setup_crop_bindings(self):
        """Set up mouse bindings for manual crop (called after UI is created)."""
        self.preview_canvas.bind("<Button-1>", self._crop_start)
        self.preview_canvas.bind("<B1-Motion>", self._crop_drag)
        self.preview_canvas.bind("<ButtonRelease-1>", self._crop_end)
    
    def _refresh_ports(self):
        """Refresh available COM ports and test connection."""
        ports = list_available_ports()
        port_names = [p['device'] for p in ports]
        
        if port_names:
            self.port_dropdown.configure(values=port_names)
            if self.selected_port.get() not in port_names:
                self.selected_port.set(port_names[0])
            
            # Test connection to selected port
            self._test_connection()
        else:
            self.port_dropdown.configure(values=["No ports found"])
            self.status_label.configure(text="● No ports", text_color="#ff6b6b")
        
        self._set_status(f"Found {len(port_names)} COM port(s)")
    
    def _test_connection(self):
        """Test connection to selected port."""
        port = self.selected_port.get()
        if not port or port == "No ports found":
            self.status_label.configure(text="● Disconnected", text_color="#ff6b6b")
            return
        
        try:
            printer = BluetoothPrinter(port=port, baudrate=9600)
            if printer.connect():
                self.status_label.configure(text="● Connected", text_color="#4ade80")
                self._set_status(f"Connected to {port}")
                printer.disconnect()
            else:
                self.status_label.configure(text="● Disconnected", text_color="#ff6b6b")
                self._set_status(f"Could not connect to {port}")
        except Exception:
            self.status_label.configure(text="● Disconnected", text_color="#ff6b6b")
    
    def _set_size(self, width: int, height: int):
        """Set label size from preset buttons."""
        self.label_width.set(width)
        self.label_height.set(height)
        
        # Update button states
        presets = [
            (self.size_4x6_btn, 101, 152),
            (self.size_4x4_btn, 101, 101),
            (self.size_2x1_btn, 50, 25),
        ]
        
        is_preset = False
        for btn, w, h in presets:
            if w == width and h == height:
                btn.configure(fg_color="#2d5a27", border_width=0)
                is_preset = True
            else:
                btn.configure(fg_color="transparent", border_width=1)
        
        # Update custom button
        if is_preset:
            self.size_custom_btn.configure(fg_color="transparent", border_width=1)
        else:
            self.size_custom_btn.configure(fg_color="#2d5a27", border_width=0)
        
        self._update_preview()
    
    def _set_custom_size(self):
        """Activate custom size mode - uses values from width/height entries."""
        # Deselect all preset buttons
        for btn in [self.size_4x6_btn, self.size_4x4_btn, self.size_2x1_btn]:
            btn.configure(fg_color="transparent", border_width=1)
        self.size_custom_btn.configure(fg_color="#2d5a27", border_width=0)
        
        self._set_status("Enter custom width/height in mm below")
        self._update_preview()
    
    def _on_custom_size_change(self):
        """Called when custom size entries are modified."""
        # Check if current size matches a preset
        w, h = self.label_width.get(), self.label_height.get()
        presets = [(101, 152), (101, 101), (50, 25)]
        
        if (w, h) not in presets:
            # Switch to custom mode
            self._set_custom_size()
        else:
            # Update to matching preset
            self._set_size(w, h)
    
    def _select_pdf(self):
        """Open file dialog to select PDF."""
        file_path = filedialog.askopenfilename(
            title="Select PDF Label",
            filetypes=[
                ("PDF files", "*.pdf"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.pdf_path = Path(file_path)
            self.filename_label.configure(text=self.pdf_path.name)
            self._set_status(f"Loaded: {self.pdf_path.name}")
            self._update_preview()
    
    def _update_preview(self):
        """Update the preview image."""
        if not self.pdf_path or not self.pdf_path.exists():
            return
        
        try:
            # If manual crop mode and no crop applied yet, show full PDF for selection
            if self.manual_crop.get() and not self.crop_applied:
                self._show_full_pdf_for_crop()
                return
            
            # Otherwise show the final processed preview (what will print)
            self._show_processed_preview()
            
        except Exception as e:
            logger.error(f"Preview error: {e}")
            import traceback
            traceback.print_exc()
            self._set_status(f"Preview error: {e}")
    
    def _show_full_pdf_for_crop(self):
        """Show the full PDF for manual crop selection."""
        from bt_printer import BluetoothPrinter
        
        pdf_doc = fitz.open(self.pdf_path)
        page = pdf_doc[0]
        
        # Render at 150 DPI for selection (good balance of quality and speed)
        render_dpi = 150
        zoom = render_dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        img_data = pix.tobytes("ppm")
        image = Image.open(BytesIO(img_data))
        pdf_doc.close()
        
        # Store original PDF image for crop coordinate calculation
        self.original_pdf_image = image.copy()
        self.original_pdf_size = image.size
        
        # Scale for display
        canvas_width = 400
        canvas_height = 600
        
        preview_image = image.copy()
        preview_image.thumbnail((canvas_width, canvas_height), Image.Resampling.LANCZOS)
        
        # Store scale factors for coordinate conversion
        self.pdf_display_scale_x = preview_image.width / image.width
        self.pdf_display_scale_y = preview_image.height / image.height
        
        # Convert to PhotoImage
        from PIL import ImageTk
        photo = ImageTk.PhotoImage(preview_image)
        
        # Clear canvas
        self.preview_canvas.delete("all")
        self.canvas_placeholder.grid_remove()
        
        # Store preview size for coordinate conversion
        self.preview_display_size = preview_image.size
        
        # Center image on canvas
        x = (canvas_width - preview_image.width) // 2
        y = (canvas_height - preview_image.height) // 2
        self.preview_offset = (x, y)
        
        self.canvas_image_id = self.preview_canvas.create_image(
            x + preview_image.width // 2,
            y + preview_image.height // 2,
            image=photo,
            anchor="center"
        )
        
        self.preview_canvas.image = photo
        self.preview_canvas.photo = photo
        
        # Draw pending selection rectangle if exists
        if self.pending_crop:
            self._draw_pending_crop()
        
        self._set_status("Draw a rectangle to select the area you want to print")
    
    def _show_processed_preview(self):
        """Show the processed preview (what will actually print)."""
        from bt_printer import BluetoothPrinter
        
        dummy_printer = BluetoothPrinter()
        
        pdf_doc = fitz.open(self.pdf_path)
        page = pdf_doc[0]
        
        render_dpi = 300
        zoom = render_dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        img_data = pix.tobytes("ppm")
        image = Image.open(BytesIO(img_data))
        pdf_doc.close()
        
        if image.mode != 'L':
            image = image.convert('L')
        
        # Apply crop
        if self.manual_crop.get() and self.crop_applied and self.crop_coords:
            # Scale crop coords from 150 DPI to 300 DPI
            scale = 300 / 150
            left, top, right, bottom = self.crop_coords
            left = int(left * scale)
            top = int(top * scale)
            right = int(right * scale)
            bottom = int(bottom * scale)
            
            # Clamp to image bounds
            left = max(0, min(left, image.width))
            right = max(0, min(right, image.width))
            top = max(0, min(top, image.height))
            bottom = max(0, min(bottom, image.height))
            
            if right > left and bottom > top:
                image = image.crop((left, top, right, bottom))
                logger.info(f"Applied crop: ({left}, {top}) to ({right}, {bottom})")
        elif self.auto_crop.get():
            image = dummy_printer._crop_to_label(image)
        
        # Target size
        dpi = 203
        target_width_dots = int(self.label_width.get() * dpi / 25.4)
        target_height_dots = int(self.label_height.get() * dpi / 25.4)
        
        # Auto-rotate
        img_w, img_h = image.size
        if self.auto_rotate.get():
            img_is_landscape = img_w > img_h
            label_is_landscape = target_width_dots > target_height_dots
            
            if img_is_landscape != label_is_landscape:
                image = image.rotate(90, expand=True)
                img_w, img_h = image.size
        
        # Flip vertical (rotate 180 degrees to flip without mirroring)
        if self.flip_vertical.get():
            # Rotate 180 degrees to flip top-to-bottom without mirroring text
            image = image.rotate(180, expand=False)
        
        # Scale
        scale_w = target_width_dots / img_w
        scale_h = target_height_dots / img_h
        scale = min(scale_w, scale_h)
        
        new_width = int(img_w * scale)
        new_height = int(img_h * scale)
        
        image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
        
        # Center on label
        final_image = Image.new('L', (target_width_dots, target_height_dots), 255)
        x_offset = (target_width_dots - new_width) // 2
        y_offset = (target_height_dots - new_height) // 2
        final_image.paste(image, (x_offset, y_offset))
        
        # Convert to 1-bit
        final_image = final_image.point(lambda x: 0 if x < 128 else 255, '1')
        
        # Store non-inverted version for printing
        self.cropped_image = final_image.copy()
        
        # Preview always shows normal (non-inverted) - what the label should look like
        # We'll invert before sending to printer if checkbox is OFF
        preview_image = final_image.copy()
        
        # Display
        canvas_width = 400
        canvas_height = 600
        
        preview_image = preview_image.copy()
        preview_image.thumbnail((canvas_width, canvas_height), Image.Resampling.LANCZOS)
        
        from PIL import ImageTk
        photo = ImageTk.PhotoImage(preview_image)
        
        self.preview_canvas.delete("all")
        self.canvas_placeholder.grid_remove()
        
        # Resize canvas to match image exactly
        canvas_width = preview_image.width
        canvas_height = preview_image.height
        self.preview_canvas.configure(width=canvas_width, height=canvas_height)
        
        # Center image on canvas (canvas is already centered in frame via grid)
        x = canvas_width // 2
        y = canvas_height // 2
        
        self.canvas_image_id = self.preview_canvas.create_image(
            x, y,
            image=photo,
            anchor="center"
        )
        
        self.preview_canvas.image = photo
        self.preview_canvas.photo = photo
        
        self._set_status("Preview shows what will print")
    
    def _toggle_manual_crop(self):
        """Toggle manual crop mode."""
        if self.manual_crop.get():
            self.auto_crop.set(False)
            self.crop_applied = False
            self.crop_coords = None
            self.pending_crop = None
            self.apply_crop_btn.configure(state="disabled")
            self.clear_crop_btn.configure(state="normal")
            self._set_status("Draw a rectangle on the preview to select print area")
        else:
            self.crop_applied = False
            self.crop_coords = None
            self.pending_crop = None
            self.apply_crop_btn.configure(state="disabled")
            self.clear_crop_btn.configure(state="disabled")
        self._update_preview()
    
    def _crop_start(self, event):
        """Start manual crop selection."""
        if not self.manual_crop.get() or self.crop_applied:
            return
        if not hasattr(self, 'original_pdf_size') or not self.original_pdf_size:
            return
        
        self.crop_start = (event.x, event.y)
        # Delete any existing rectangle
        self.preview_canvas.delete("pending_crop")
    
    def _crop_drag(self, event):
        """Update crop rectangle while dragging."""
        if not self.manual_crop.get() or self.crop_applied:
            return
        if not self.crop_start:
            return
        
        # Delete old rectangle
        self.preview_canvas.delete("pending_crop")
        
        # Draw new rectangle
        x1, y1 = self.crop_start
        x2, y2 = event.x, event.y
        
        self.preview_canvas.create_rectangle(
            x1, y1, x2, y2,
            outline="#4ade80",
            width=3,
            tags="pending_crop"
        )
    
    def _crop_end(self, event):
        """Finish manual crop selection."""
        if not self.manual_crop.get() or self.crop_applied:
            return
        if not self.crop_start:
            return
        if not hasattr(self, 'original_pdf_size') or not self.original_pdf_size:
            return
        
        x1, y1 = self.crop_start
        x2, y2 = event.x, event.y
        
        # Normalize coordinates
        canvas_left = min(x1, x2)
        canvas_right = max(x1, x2)
        canvas_top = min(y1, y2)
        canvas_bottom = max(y1, y2)
        
        # Convert canvas coordinates to original PDF image coordinates
        if self.canvas_image_id and hasattr(self, 'preview_offset'):
            bbox = self.preview_canvas.bbox(self.canvas_image_id)
            if bbox:
                img_x1, img_y1, img_x2, img_y2 = bbox
                
                # Convert canvas coords to PDF coords
                pdf_left = (canvas_left - img_x1) / self.pdf_display_scale_x
                pdf_right = (canvas_right - img_x1) / self.pdf_display_scale_x
                pdf_top = (canvas_top - img_y1) / self.pdf_display_scale_y
                pdf_bottom = (canvas_bottom - img_y1) / self.pdf_display_scale_y
                
                # Clamp to image bounds
                pdf_left = max(0, min(pdf_left, self.original_pdf_size[0]))
                pdf_right = max(0, min(pdf_right, self.original_pdf_size[0]))
                pdf_top = max(0, min(pdf_top, self.original_pdf_size[1]))
                pdf_bottom = max(0, min(pdf_bottom, self.original_pdf_size[1]))
                
                if abs(pdf_right - pdf_left) > 20 and abs(pdf_bottom - pdf_top) > 20:
                    self.pending_crop = (int(pdf_left), int(pdf_top), int(pdf_right), int(pdf_bottom))
                    self.apply_crop_btn.configure(state="normal")
                    self._set_status(f"Selection: {int(pdf_right - pdf_left)}×{int(pdf_bottom - pdf_top)} - Click 'Apply Crop'")
        
        self.crop_start = None
    
    def _draw_pending_crop(self):
        """Draw the pending crop rectangle."""
        if not self.pending_crop or not hasattr(self, 'original_pdf_size'):
            return
        
        left, top, right, bottom = self.pending_crop
        
        if self.canvas_image_id:
            bbox = self.preview_canvas.bbox(self.canvas_image_id)
            if bbox:
                img_x1, img_y1, img_x2, img_y2 = bbox
                
                # Convert PDF coords to canvas coords
                canvas_left = left * self.pdf_display_scale_x + img_x1
                canvas_right = right * self.pdf_display_scale_x + img_x1
                canvas_top = top * self.pdf_display_scale_y + img_y1
                canvas_bottom = bottom * self.pdf_display_scale_y + img_y1
                
                self.preview_canvas.create_rectangle(
                    canvas_left, canvas_top, canvas_right, canvas_bottom,
                    outline="#4ade80",
                    width=3,
                    tags="pending_crop"
                )
    
    def _apply_crop(self):
        """Apply the pending crop selection."""
        if self.pending_crop:
            self.crop_coords = self.pending_crop
            self.crop_applied = True
            self.pending_crop = None
            self.apply_crop_btn.configure(state="disabled")
            self._set_status("Crop applied - preview shows what will print")
            self._update_preview()
    
    def _clear_crop(self):
        """Clear manual crop selection."""
        self.crop_coords = None
        self.pending_crop = None
        self.crop_applied = False
        self.preview_canvas.delete("pending_crop")
        self.apply_crop_btn.configure(state="disabled")
        self._update_preview()
        self._set_status("Crop cleared - draw a new selection")
    
    def _print_label(self):
        """Print the label."""
        if not self.pdf_path:
            messagebox.showwarning("No File", "Please select a PDF file first.")
            return
        
        # Disable button and show progress
        self.print_btn.configure(state="disabled", text="Printing...")
        self.progress.grid()
        self.progress.set(0.2)
        
        # Print in background thread
        thread = threading.Thread(target=self._do_print)
        thread.start()
    
    def _do_print(self):
        """Actual print operation (runs in background thread)."""
        try:
            self._set_status("Connecting to printer...")
            self.after(0, lambda: self.progress.set(0.3))
            
            printer = BluetoothPrinter(
                port=self.selected_port.get(),
                baudrate=9600
            )
            
            if not printer.connect():
                self.after(0, lambda: messagebox.showerror(
                    "Connection Failed",
                    f"Could not connect to {self.selected_port.get()}.\n\nMake sure the printer is on and paired."
                ))
                return
            
            self.after(0, lambda: self.progress.set(0.5))
            self._set_status("Sending label to printer...")
            
            try:
                # If we have a processed image (from manual crop), use it directly
                # Note: invert_colors OFF (False) = invert for printing, ON (True) = don't invert
                # cropped_image is stored non-inverted, so we invert it here if needed
                if self.manual_crop.get() and self.crop_applied and hasattr(self, 'cropped_image') and self.cropped_image:
                    print_image = self.cropped_image.copy()
                    if not self.invert_colors.get():  # OFF = invert before sending
                        from PIL import ImageOps
                        print_image = ImageOps.invert(print_image.convert('L')).convert('1')
                    success = printer.print_image_tspl(
                        print_image,
                        label_width_mm=self.label_width.get(),
                        label_height_mm=self.label_height.get()
                    )
                else:
                    # Use normal PDF processing for auto-crop mode
                    # Note: invert_colors OFF (False) = invert for printing, ON (True) = don't invert
                    success = printer.print_pdf(
                        self.pdf_path,
                        label_width_mm=self.label_width.get(),
                        label_height_mm=self.label_height.get(),
                        auto_crop=self.auto_crop.get(),
                        auto_rotate=self.auto_rotate.get(),
                        flip_vertical=self.flip_vertical.get(),
                        invert=not self.invert_colors.get()  # Invert the checkbox value
                    )
                
                self.after(0, lambda: self.progress.set(1.0))
                
                if success:
                    self._set_status("Label printed successfully!")
                    self.after(0, lambda: self.status_label.configure(
                        text="● Connected",
                        text_color="#4ade80"
                    ))
                else:
                    self._set_status("Print may have failed - check printer")
                    
            finally:
                printer.disconnect()
                
        except Exception as e:
            logger.error(f"Print error: {e}")
            self._set_status(f"Error: {e}")
            self.after(0, lambda: messagebox.showerror("Print Error", str(e)))
        
        finally:
            # Re-enable button
            self.after(0, self._reset_print_button)
    
    def _reset_print_button(self):
        """Reset print button state."""
        self.print_btn.configure(state="normal", text="🖨️ Print Label")
        self.progress.grid_remove()
    
    def _set_status(self, text: str):
        """Update status text (thread-safe)."""
        self.after(0, lambda: self.status_text.configure(text=text))
    
    def _show_about(self):
        """Show About dialog."""
        about_text = f"""Bluetooth Label Printer
Version {__version__}

A simple tool to print shipping labels to Bluetooth thermal printers.

Supports MVGGES, TSC, Xprinter, and most TSPL-compatible label printers.

GitHub: github.com/elitescouter/PrintToBTLabel

MIT License"""
        
        messagebox.showinfo("About", about_text)


def main():
    app = LabelPrinterApp()
    app.mainloop()


if __name__ == "__main__":
    main()


