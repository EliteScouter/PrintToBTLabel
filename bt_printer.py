"""
Bluetooth COM Port Printer Bridge
Sends print data to a Bluetooth COM port (COM5) for label printing.
"""
import logging
import serial
import serial.tools.list_ports
from typing import Optional, Union
from pathlib import Path
import time

# Configure logging first
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Try to import PDF/image support
try:
    from PIL import Image
    import fitz  # PyMuPDF - easier than pdf2image (no poppler needed)
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False
    logger.warning("PDF support not available. Install PyMuPDF and Pillow: pip install PyMuPDF Pillow")



class BluetoothPrinter:
    """Handles printing to a Bluetooth COM port."""
    
    def __init__(
        self,
        port: str = "COM5",
        baudrate: int = 9600,
        timeout: float = 5.0,
        bytesize: int = serial.EIGHTBITS,
        parity: str = serial.PARITY_NONE,
        stopbits: int = serial.STOPBITS_ONE
    ):
        """
        Initialize the Bluetooth printer connection.
        
        Args:
            port: COM port name (default: COM5)
            baudrate: Baud rate (default: 9600, common for label printers)
            timeout: Read timeout in seconds
            bytesize: Number of data bits
            parity: Parity checking
            stopbits: Number of stop bits
        """
        self.port = port
        self.baudrate = baudrate
        self.timeout = timeout
        self.bytesize = bytesize
        self.parity = parity
        self.stopbits = stopbits
        self.connection: Optional[serial.Serial] = None
        
    def connect(self) -> bool:
        """
        Establish connection to the Bluetooth COM port.
        
        Returns:
            True if connection successful, False otherwise
        """
        if self.connection and self.connection.is_open:
            logger.info(f"Already connected to {self.port}")
            return True
            
        try:
            logger.info(f"Connecting to {self.port} at {self.baudrate} baud...")
            self.connection = serial.Serial(
                port=self.port,
                baudrate=self.baudrate,
                timeout=self.timeout,
                bytesize=self.bytesize,
                parity=self.parity,
                stopbits=self.stopbits
            )
            
            # Give the connection a moment to stabilize
            time.sleep(0.5)
            
            if self.connection.is_open:
                logger.info(f"Successfully connected to {self.port}")
                return True
            else:
                logger.error(f"Failed to open connection to {self.port}")
                return False
                
        except serial.SerialException as e:
            logger.error(f"Serial connection error: {e}")
            return False
        except Exception as e:
            logger.error(f"Unexpected error connecting to {self.port}: {e}")
            return False
    
    def disconnect(self) -> None:
        """Close the connection to the printer."""
        if self.connection and self.connection.is_open:
            self.connection.close()
            logger.info(f"Disconnected from {self.port}")
            self.connection = None
    
    def is_connected(self) -> bool:
        """Check if printer is connected."""
        return self.connection is not None and self.connection.is_open
    
    def send_raw(self, data: Union[bytes, bytearray], debug: bool = False) -> bool:
        """
        Send raw bytes to the printer.
        
        Args:
            data: Raw bytes to send
            debug: If True, log hex representation of data
            
        Returns:
            True if sent successfully, False otherwise
        """
        if not self.is_connected():
            logger.error("Printer not connected. Call connect() first.")
            return False
            
        try:
            if debug:
                hex_str = ' '.join(f'{b:02X}' for b in data)
                logger.debug(f"Sending bytes (hex): {hex_str}")
                logger.debug(f"Sending bytes (repr): {repr(data)}")
            
            bytes_written = self.connection.write(data)
            self.connection.flush()  # Ensure data is sent immediately
            logger.info(f"Sent {bytes_written} bytes to printer")
            return True
        except serial.SerialException as e:
            logger.error(f"Error sending data: {e}")
            return False
        except Exception as e:
            logger.error(f"Unexpected error sending data: {e}")
            return False
    
    def send_text(self, text: str, encoding: str = "utf-8") -> bool:
        """
        Send text to the printer.
        
        Args:
            text: Text string to print
            encoding: Text encoding (default: utf-8)
            
        Returns:
            True if sent successfully, False otherwise
        """
        try:
            data = text.encode(encoding)
            return self.send_raw(data)
        except UnicodeEncodeError as e:
            logger.error(f"Encoding error: {e}")
            return False
    
    def send_file(self, file_path: Union[str, Path]) -> bool:
        """
        Send file contents to the printer.
        
        Args:
            file_path: Path to file to print
            
        Returns:
            True if sent successfully, False otherwise
        """
        file_path = Path(file_path)
        
        if not file_path.exists():
            logger.error(f"File not found: {file_path}")
            return False
            
        try:
            with open(file_path, 'rb') as f:
                data = f.read()
            logger.info(f"Reading {len(data)} bytes from {file_path}")
            return self.send_raw(data)
        except IOError as e:
            logger.error(f"Error reading file: {e}")
            return False
    
    def send_esc_pos_command(self, command: bytes) -> bool:
        """
        Send ESC/POS command (common for label printers).
        
        Args:
            command: ESC/POS command bytes
            
        Returns:
            True if sent successfully, False otherwise
        """
        return self.send_raw(command)
    
    def initialize_printer(self) -> bool:
        """
        Initialize printer with common ESC/POS commands.
        This is a generic initialization - adjust for your specific printer.
        """
        if not self.is_connected():
            return False
            
        try:
            # ESC @ - Initialize printer
            init_cmd = b'\x1B\x40'
            self.send_raw(init_cmd)
            time.sleep(0.1)
            
            # Set character encoding (UTF-8)
            encoding_cmd = b'\x1B\x74\x10'  # ESC t 16 (UTF-8)
            self.send_raw(encoding_cmd)
            time.sleep(0.1)
            
            logger.info("Printer initialized")
            return True
        except Exception as e:
            logger.error(f"Error initializing printer: {e}")
            return False
    
    def print_label(
        self,
        text: str,
        cut_after: bool = True,
        feed_lines: int = 3,
        simple_mode: bool = False,
        use_tspl: bool = True,
        label_width_mm: int = 40,
        label_height_mm: int = 30,
        gap_mm: int = 2
    ) -> bool:
        """
        Print a label with text.
        
        Args:
            text: Text to print on label
            cut_after: Whether to cut after printing (if supported)
            feed_lines: Number of lines to feed after printing
            simple_mode: If True, send plain text without initialization
            use_tspl: Use TSPL protocol (default: True for label printers)
            label_width_mm: Label width in millimeters
            label_height_mm: Label height in millimeters
            gap_mm: Gap between labels in millimeters
            
        Returns:
            True if successful, False otherwise
        """
        if not self.is_connected():
            logger.error("Printer not connected")
            return False
            
        try:
            if use_tspl:
                return self.print_label_tspl(
                    text, cut_after, label_width_mm, label_height_mm, gap_mm
                )
            
            # Legacy ESC/POS mode
            if not simple_mode:
                self.initialize_printer()
            
            if not self.send_text(text + "\n"):
                return False
            
            for _ in range(feed_lines):
                self.send_text("\n")
            
            if cut_after and not simple_mode:
                cut_cmd = b'\x1D\x56\x00'
                self.send_raw(cut_cmd)
            
            logger.info("Label printed successfully")
            return True
        except Exception as e:
            logger.error(f"Error printing label: {e}")
            return False
    
    def print_label_tspl(
        self,
        text: str,
        cut_after: bool = True,
        label_width_mm: int = 40,
        label_height_mm: int = 30,
        gap_mm: int = 2,
        font: str = "2",
        x: int = 10,
        y: int = 10
    ) -> bool:
        """
        Print label using TSPL (TSC Printer Language) protocol.
        This is the standard for Chinese label printers like MVGGES.
        
        Args:
            text: Text to print
            cut_after: Whether to cut after printing
            label_width_mm: Label width in millimeters
            label_height_mm: Label height in millimeters
            gap_mm: Gap between labels in millimeters
            font: Font size (0-8, where 2 is common)
            x: X position in dots (default: 10)
            y: Y position in dots (default: 10)
            
        Returns:
            True if successful, False otherwise
        """
        if not self.is_connected():
            logger.error("Printer not connected")
            return False
        
        try:
            # Build TSPL command sequence
            # SIZE: width,height (in mm or dots)
            # GAP: gap,distance (in mm)
            # DIRECTION: 0=normal, 1=90°, 2=180°, 3=270°
            # CLS: Clear print buffer
            # TEXT: x,y,font,rotation,x-multi,y-multi,text
            # PRINT: quantity
            
            lines = text.split('\n')
            
            # Build TSPL command - match the format that worked in test_tspl.py
            # Use newlines (not \r\n) as that's what worked
            tspl_cmd = f"SIZE {label_width_mm} mm, {label_height_mm} mm\n"
            tspl_cmd += f"GAP {gap_mm} mm, 0 mm\n"
            tspl_cmd += "DIRECTION 1\n"
            tspl_cmd += "CLS\n"
            
            # Add each line of text
            current_y = y
            for line in lines:
                if line.strip():  # Skip empty lines
                    # Escape quotes in text
                    escaped_text = line.replace('"', '\\"')
                    tspl_cmd += f'TEXT {x},{current_y},"{font}",0,1,1,"{escaped_text}"\n'
                    current_y += 40  # Move down for next line (40 dots spacing)
            
            # Print command
            tspl_cmd += "PRINT 1\n"
            tspl_cmd += "\n"  # Extra newline at end (as in working test)
            
            logger.debug(f"TSPL command:\n{tspl_cmd}")
            logger.info("Sending TSPL label print command")
            
            # Send the command
            return self.send_text(tspl_cmd, encoding="utf-8")
            
        except Exception as e:
            logger.error(f"Error printing TSPL label: {e}")
            return False
    
    def print_pdf(
        self,
        pdf_path: Union[str, Path],
        label_width_mm: int = 101,
        label_height_mm: int = 152,
        gap_mm: int = 2,
        dpi: int = 203,
        auto_rotate: bool = True,
        auto_crop: bool = True,
        invert: bool = True,
        manual_crop_coords: Optional[tuple] = None,
        flip_vertical: bool = False
    ) -> bool:
        """
        Print a PDF file as a label using TSPL BITMAP command.
        Common shipping label sizes:
        - 4x6 inches = 101.6 x 152.4 mm (default)
        - 4x4 inches = 101.6 x 101.6 mm
        
        Args:
            pdf_path: Path to PDF file
            label_width_mm: Label width in millimeters (default: 101 = 4 inches)
            label_height_mm: Label height in millimeters (default: 152 = 6 inches)
            gap_mm: Gap between labels in millimeters
            dpi: Printer DPI (default: 203, common for label printers)
            auto_rotate: Automatically rotate if needed to fit label
            auto_crop: Automatically crop whitespace from PDF
            invert: Invert colors (default: True for correct printing)
            
        Returns:
            True if successful, False otherwise
        """
        if not PDF_SUPPORT:
            logger.error("PDF support not available. Install: pip install PyMuPDF Pillow")
            return False
        
        if not self.is_connected():
            logger.error("Printer not connected")
            return False
        
        pdf_path = Path(pdf_path)
        if not pdf_path.exists():
            logger.error(f"PDF file not found: {pdf_path}")
            return False
        
        try:
            logger.info(f"Converting PDF to image: {pdf_path}")
            
            # Render PDF at higher resolution for quality, then resize
            render_dpi = 300  # Render at 300 DPI for quality
            
            pdf_doc = fitz.open(pdf_path)
            
            if len(pdf_doc) == 0:
                logger.error("PDF has no pages")
                pdf_doc.close()
                return False
            
            page = pdf_doc[0]
            
            # Render at high DPI
            zoom = render_dpi / 72.0
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            
            # Convert to PIL Image
            from io import BytesIO
            img_data = pix.tobytes("ppm")
            image = Image.open(BytesIO(img_data))
            
            pdf_doc.close()
            
            logger.info(f"PDF rendered: {image.size[0]}x{image.size[1]} pixels")
            
            # Convert to grayscale for processing
            if image.mode != 'L':
                image = image.convert('L')
            
            # Manual crop takes precedence
            # Coords are in 150 DPI space (from preview), scale to 300 DPI
            if manual_crop_coords:
                scale_factor = render_dpi / 150.0  # 300/150 = 2.0
                left, top, right, bottom = manual_crop_coords
                left = int(left * scale_factor)
                top = int(top * scale_factor)
                right = int(right * scale_factor)
                bottom = int(bottom * scale_factor)
                
                # Clamp to image bounds
                left = max(0, min(left, image.width))
                right = max(0, min(right, image.width))
                top = max(0, min(top, image.height))
                bottom = max(0, min(bottom, image.height))
                
                if right > left and bottom > top:
                    image = image.crop((left, top, right, bottom))
                    logger.info(f"Manual crop applied: ({left}, {top}) to ({right}, {bottom}) at {render_dpi} DPI")
            # Auto-crop to label area (ignores text outside label)
            elif auto_crop:
                image = self._crop_to_label(image)
                logger.info(f"After crop: {image.size[0]}x{image.size[1]} pixels")
            
            # Calculate target size in dots at printer DPI
            target_width_dots = int(label_width_mm * dpi / 25.4)
            target_height_dots = int(label_height_mm * dpi / 25.4)
            
            logger.info(f"Target label size: {target_width_dots}x{target_height_dots} dots")
            
            # Auto-rotate if needed
            img_w, img_h = image.size
            img_aspect = img_w / img_h
            label_aspect = target_width_dots / target_height_dots
            
            # If image is landscape and label is portrait (or vice versa), rotate
            if auto_rotate:
                img_is_landscape = img_w > img_h
                label_is_landscape = target_width_dots > target_height_dots
                
                if img_is_landscape != label_is_landscape:
                    logger.info("Rotating image 90 degrees to fit label orientation")
                    image = image.rotate(90, expand=True)
                    img_w, img_h = image.size
            
            # Flip vertical if requested (rotate 180 degrees to flip without mirroring)
            if flip_vertical:
                # Rotate 180 degrees to flip top-to-bottom without mirroring text
                image = image.rotate(180, expand=False)
            
            # Scale to fit label while maintaining aspect ratio
            scale_w = target_width_dots / img_w
            scale_h = target_height_dots / img_h
            scale = min(scale_w, scale_h)  # Fit within label
            
            new_width = int(img_w * scale)
            new_height = int(img_h * scale)
            
            logger.info(f"Scaling image to: {new_width}x{new_height} dots")
            
            # Use high-quality resize
            image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            # Apply manual crop to final_image if provided
            if manual_crop_coords:
                # Crop coords are in final_image space (target_width_dots x target_height_dots)
                # Scale the image first, then crop from final_image
                left, top, right, bottom = manual_crop_coords
                # Clamp to final_image bounds
                left = max(0, min(int(left), target_width_dots))
                right = max(0, min(int(right), target_width_dots))
                top = max(0, min(int(top), target_height_dots))
                bottom = max(0, min(int(bottom), target_height_dots))
                
                if right > left and bottom > top:
                    # Create temp final_image to crop from
                    temp_final = Image.new('L', (target_width_dots, target_height_dots), 255)
                    temp_final.paste(image, ((target_width_dots - new_width) // 2, 
                                           (target_height_dots - new_height) // 2))
                    # Crop the region
                    cropped_region = temp_final.crop((left, top, right, bottom))
                    # Resize back to fit label
                    cropped_region = cropped_region.resize((target_width_dots, target_height_dots), 
                                                          Image.Resampling.LANCZOS)
                    final_image = cropped_region
                    logger.info(f"Manual crop applied to final image: ({left}, {top}) to ({right}, {bottom})")
                else:
                    # Invalid crop, use normal flow
                    final_image = Image.new('L', (target_width_dots, target_height_dots), 255)
                    x_offset = (target_width_dots - new_width) // 2
                    y_offset = (target_height_dots - new_height) // 2
                    final_image.paste(image, (x_offset, y_offset))
            else:
                # Create final image centered on label
                final_image = Image.new('L', (target_width_dots, target_height_dots), 255)  # White background
                x_offset = (target_width_dots - new_width) // 2
                y_offset = (target_height_dots - new_height) // 2
                final_image.paste(image, (x_offset, y_offset))
            
            # Convert to 1-bit AFTER scaling for better quality
            # Use dithering for better barcode reproduction
            final_image = final_image.point(lambda x: 0 if x < 128 else 255, '1')
            
            if invert:
                from PIL import ImageOps
                final_image = ImageOps.invert(final_image.convert('L')).convert('1')
            
            # Print the image
            return self.print_image_tspl(
                final_image,
                label_width_mm=label_width_mm,
                label_height_mm=label_height_mm,
                gap_mm=gap_mm
            )
            
        except Exception as e:
            logger.error(f"Error printing PDF: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _crop_whitespace(self, image: Image.Image, threshold: int = 250) -> Image.Image:
        """
        Crop whitespace from image edges.
        
        Args:
            image: PIL Image in grayscale mode
            threshold: Pixels lighter than this are considered white (0-255)
            
        Returns:
            Cropped image
        """
        # Convert to grayscale if needed
        if image.mode != 'L':
            image = image.convert('L')
        
        # Create binary image where white pixels are 0, non-white are 255
        # This helps getbbox() find the content area
        binary = image.point(lambda x: 0 if x > threshold else 255)
        
        # Get bounding box of non-white content
        bbox = binary.getbbox()
        
        if bbox:
            # Add small margin (10 pixels)
            margin = 10
            left = max(0, bbox[0] - margin)
            top = max(0, bbox[1] - margin)
            right = min(image.width, bbox[2] + margin)
            bottom = min(image.height, bbox[3] + margin)
            
            logger.info(f"Cropping from ({bbox[0]}, {bbox[1]}) to ({bbox[2]}, {bbox[3]})")
            return image.crop((left, top, right, bottom))
        
        return image
    
    def _find_label_boundary(self, image: Image.Image) -> Optional[tuple]:
        """
        Find the shipping label boundary (rectangular border).
        Looks for the largest rectangle in the image.
        
        Args:
            image: PIL Image in grayscale mode
            
        Returns:
            Bounding box (left, top, right, bottom) or None if not found
        """
        if image.mode != 'L':
            image = image.convert('L')
        
        width, height = image.size
        pixels = image.load()
        
        # Threshold to find dark pixels (potential border)
        threshold = 128
        
        # Scan for the label rectangle
        # Find first significant horizontal line from top
        top = 0
        for y in range(height):
            dark_count = sum(1 for x in range(width) if pixels[x, y] < threshold)
            # If more than 30% of the row is dark, likely a border
            if dark_count > width * 0.3:
                top = y
                break
        
        # Find last significant horizontal line from bottom
        bottom = height
        for y in range(height - 1, -1, -1):
            dark_count = sum(1 for x in range(width) if pixels[x, y] < threshold)
            if dark_count > width * 0.3:
                bottom = y
                break
        
        # Find first significant vertical line from left
        left = 0
        for x in range(width):
            dark_count = sum(1 for y in range(height) if pixels[x, y] < threshold)
            if dark_count > height * 0.3:
                left = x
                break
        
        # Find last significant vertical line from right
        right = width
        for x in range(width - 1, -1, -1):
            dark_count = sum(1 for y in range(height) if pixels[x, y] < threshold)
            if dark_count > height * 0.3:
                right = x
                break
        
        # Validate the found rectangle
        if right > left + 100 and bottom > top + 100:  # Minimum size
            return (left, top, right, bottom)
        
        return None
    
    def _crop_to_label(self, image: Image.Image) -> Image.Image:
        """
        Crop image to just the shipping label area inside the dashed/solid border.
        Looks for the rectangular border that surrounds the label.
        
        Args:
            image: PIL Image
            
        Returns:
            Cropped image containing just the label
        """
        if image.mode != 'L':
            image = image.convert('L')
        
        width, height = image.size
        pixels = image.load()
        
        # Strategy: Find horizontal lines that span most of the width (border lines)
        # These indicate the top and bottom of the label
        
        threshold = 180  # Pixels darker than this are "dark"
        min_line_density = 0.15  # Line must have at least 15% dark pixels to be a border
        
        # Find potential horizontal border lines
        horizontal_lines = []
        for y in range(height):
            dark_count = 0
            for x in range(width):
                if pixels[x, y] < threshold:
                    dark_count += 1
            density = dark_count / width
            if density > min_line_density:
                horizontal_lines.append((y, density))
        
        # Find potential vertical border lines
        vertical_lines = []
        for x in range(width):
            dark_count = 0
            for y in range(height):
                if pixels[x, y] < threshold:
                    dark_count += 1
            density = dark_count / height
            if density > min_line_density:
                vertical_lines.append((x, density))
        
        # Find the label rectangle by looking for the border
        # Top border: first cluster of horizontal lines from top
        # Bottom border: look for a gap (the label ends before the "Return Authorization" text)
        
        label_top = 0
        label_bottom = height
        label_left = 0
        label_right = width
        
        # Find top border (first significant horizontal line)
        for y, density in horizontal_lines:
            if y > height * 0.02:  # Skip very top edge
                label_top = y
                break
        
        # Find bottom border - look for a gap in content then more content
        # This gap separates the label from text below it
        gap_threshold = 20  # Minimum gap size in pixels
        in_content = False
        gap_start = 0
        
        for y in range(label_top + 50, height):  # Start below top
            row_dark = sum(1 for x in range(width) if pixels[x, y] < threshold)
            has_content = row_dark > width * 0.02
            
            if in_content and not has_content:
                # Entering a gap
                gap_start = y
                in_content = False
            elif not in_content and has_content:
                # Exiting a gap
                gap_size = y - gap_start
                if gap_size > gap_threshold and gap_start > label_top + 100:
                    # Found significant gap - label ends before this gap
                    label_bottom = gap_start
                    break
                in_content = True
            elif y == label_top + 50:
                in_content = has_content
        
        # If no gap found, look for the dashed border line at bottom of label
        if label_bottom == height:
            # Scan from bottom up looking for a dense horizontal line (border)
            for y in range(height - 50, label_top + 100, -1):
                row_dark = sum(1 for x in range(width) if pixels[x, y] < threshold)
                if row_dark > width * 0.3:  # Strong horizontal line
                    label_bottom = y + 5
                    break
        
        # Find left and right borders within the vertical range
        for x, density in vertical_lines:
            if x > width * 0.02:
                label_left = x
                break
        
        for x, density in reversed(vertical_lines):
            if x < width * 0.98:
                label_right = x
                break
        
        # Refine: look for the actual dashed border box
        # Scan inward from edges to find where the border actually is
        
        # Refine top - find the top border line
        for y in range(min(label_top + 50, height)):
            row_dark = sum(1 for x in range(label_left, label_right) if pixels[x, y] < threshold)
            if row_dark > (label_right - label_left) * 0.2:
                label_top = y
                break
        
        # Add small inward margin to exclude the border itself
        border_margin = 3
        label_left = min(label_left + border_margin, width)
        label_top = min(label_top + border_margin, height)  
        label_right = max(label_right - border_margin, 0)
        label_bottom = max(label_bottom - border_margin, 0)
        
        logger.info(f"Found label region: ({label_left}, {label_top}) to ({label_right}, {label_bottom})")
        
        # Validate
        crop_width = label_right - label_left
        crop_height = label_bottom - label_top
        
        if crop_width > 100 and crop_height > 100:
            return image.crop((label_left, label_top, label_right, label_bottom))
        
        # Fallback
        logger.info("Label border detection failed, using whitespace crop")
        return self._crop_whitespace(image)
    
    def print_image_tspl(
        self,
        image: Image.Image,
        label_width_mm: int = 100,
        label_height_mm: int = 150,
        gap_mm: int = 2,
        x: int = 0,
        y: int = 0
    ) -> bool:
        """
        Print a PIL Image using TSPL BITMAP command.
        
        Args:
            image: PIL Image (should be mode '1' - 1-bit monochrome)
            label_width_mm: Label width in millimeters
            label_height_mm: Label height in millimeters
            gap_mm: Gap between labels in millimeters
            x: X position in dots (default: 0)
            y: Y position in dots (default: 0)
            
        Returns:
            True if successful, False otherwise
        """
        if not PDF_SUPPORT:
            logger.error("Image support not available. Install: pip install Pillow")
            return False
        
        if not self.is_connected():
            logger.error("Printer not connected")
            return False
        
        try:
            # Ensure image is 1-bit monochrome
            if image.mode != '1':
                image = image.convert('L').convert('1')
            
            width, height = image.size
            width_bytes = (width + 7) // 8  # Round up to nearest byte
            
            logger.info(f"Converting image to bitmap: {width}x{height} pixels, {width_bytes} bytes per row")
            
            # Convert image to bytes (row by row, MSB first)
            # TSPL BITMAP: 1 = black (print), 0 = white (no print)
            # PIL mode '1': 0 = black, 255 = white
            bitmap_data = bytearray()
            pixels = image.load()
            
            for row in range(height):
                for byte_col in range(width_bytes):
                    byte_val = 0
                    for bit in range(8):
                        pixel_x = byte_col * 8 + bit
                        if pixel_x < width:
                            pixel_val = pixels[pixel_x, row]
                            # Black pixel (0 in PIL) = print (1 in TSPL)
                            if pixel_val == 0:
                                byte_val |= (1 << (7 - bit))
                    bitmap_data.append(byte_val)
            
            # Build TSPL command
            tspl_cmd = f"SIZE {label_width_mm} mm, {label_height_mm} mm\n"
            tspl_cmd += f"GAP {gap_mm} mm, 0 mm\n"
            tspl_cmd += "DIRECTION 1\n"
            tspl_cmd += "CLS\n"
            
            # TSPL BITMAP command: BITMAP x,y,width_bytes,height,mode,data
            # mode: 0=overwrite, 1=OR, 2=XOR
            tspl_cmd += f"BITMAP {x},{y},{width_bytes},{height},0,"
            
            # Send command header
            self.send_text(tspl_cmd, encoding="utf-8")
            
            # Send bitmap data as raw bytes
            logger.info(f"Sending {len(bitmap_data)} bytes of bitmap data")
            self.send_raw(bytes(bitmap_data))
            
            # Send print command
            print_cmd = "PRINT 1\n\n"
            self.send_text(print_cmd, encoding="utf-8")
            
            logger.info("PDF label printed successfully")
            return True
            
        except Exception as e:
            logger.error(f"Error printing image: {e}")
            return False
    
    def print_simple(self, text: str, encoding: str = "utf-8") -> bool:
        """
        Send plain text without any initialization commands.
        Useful for testing if printer accepts simple text.
        
        Args:
            text: Text to print
            encoding: Text encoding (try 'gb2312' or 'gbk' for Chinese printers)
            
        Returns:
            True if sent successfully, False otherwise
        """
        if not self.is_connected():
            logger.error("Printer not connected")
            return False
        
        logger.info(f"Sending plain text (no initialization) with {encoding} encoding")
        return self.send_text(text + "\n\n\n", encoding=encoding)
    
    def __enter__(self):
        """Context manager entry."""
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.disconnect()


def list_available_ports() -> list:
    """List all available COM ports."""
    ports = serial.tools.list_ports.comports()
    port_list = []
    for port in ports:
        port_list.append({
            'device': port.device,
            'description': port.description,
            'manufacturer': port.manufacturer,
            'hwid': port.hwid
        })
    return port_list


def print_available_ports() -> None:
    """Print all available COM ports to console."""
    ports = list_available_ports()
    if not ports:
        print("No COM ports found.")
        return
    
    print("\nAvailable COM Ports:")
    print("-" * 60)
    for port in ports:
        print(f"Port: {port['device']}")
        print(f"  Description: {port['description']}")
        print(f"  Manufacturer: {port.get('manufacturer', 'Unknown')}")
        print(f"  HWID: {port['hwid']}")
        print()


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Print to Bluetooth COM port (COM5)"
    )
    parser.add_argument(
        "--port",
        type=str,
        default="COM5",
        help="COM port to use (default: COM5)"
    )
    parser.add_argument(
        "--baudrate",
        type=int,
        default=9600,
        help="Baud rate (default: 9600)"
    )
    parser.add_argument(
        "--text",
        type=str,
        help="Text to print"
    )
    parser.add_argument(
        "--file",
        type=str,
        help="File to print (binary)"
    )
    parser.add_argument(
        "--pdf",
        type=str,
        help="PDF file to print as label"
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=203,
        help="Printer DPI (default: 203, common for label printers)"
    )
    parser.add_argument(
        "--no-crop",
        action="store_true",
        help="Disable auto-crop of whitespace from PDF"
    )
    parser.add_argument(
        "--no-rotate",
        action="store_true",
        help="Disable auto-rotation of PDF to fit label"
    )
    parser.add_argument(
        "--no-invert",
        action="store_true",
        help="Disable color inversion (use if print is already correct)"
    )
    parser.add_argument(
        "--list-ports",
        action="store_true",
        help="List all available COM ports"
    )
    parser.add_argument(
        "--test",
        action="store_true",
        help="Send a test print"
    )
    parser.add_argument(
        "--simple",
        action="store_true",
        help="Send plain text without initialization commands"
    )
    parser.add_argument(
        "--raw-hex",
        type=str,
        help="Send raw hex bytes (e.g., '1B40' for ESC @)"
    )
    parser.add_argument(
        "--encoding",
        type=str,
        default="utf-8",
        help="Text encoding (default: utf-8, try gb2312 or gbk for Chinese printers)"
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable debug output (show hex bytes being sent)"
    )
    parser.add_argument(
        "--test-both-ports",
        action="store_true",
        help="Test both COM5 and COM6 to find the working port"
    )
    parser.add_argument(
        "--no-tspl",
        action="store_true",
        help="Disable TSPL protocol (use ESC/POS instead)"
    )
    parser.add_argument(
        "--label-width",
        type=int,
        default=101,
        help="Label width in mm (default: 101 = 4 inches)"
    )
    parser.add_argument(
        "--label-height",
        type=int,
        default=152,
        help="Label height in mm (default: 152 = 6 inches)"
    )
    
    args = parser.parse_args()
    
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    
    if args.list_ports:
        print_available_ports()
        exit(0)
    
    # Create printer instance
    printer = BluetoothPrinter(port=args.port, baudrate=args.baudrate)
    
    # Test both ports if requested
    if args.test_both_ports:
        logger.info("Testing both COM5 and COM6...")
        ports_to_test = ["COM5", "COM6"]
        
        for test_port in ports_to_test:
            logger.info(f"\n{'='*60}")
            logger.info(f"Testing {test_port}")
            logger.info(f"{'='*60}")
            
            test_printer = BluetoothPrinter(port=test_port, baudrate=args.baudrate)
            try:
                if test_printer.connect():
                    logger.info(f"Connected to {test_port} - sending test print...")
                    test_printer.print_simple("TEST PRINT\nIf you see this,\nthis port works!", encoding=args.encoding)
                    time.sleep(3)  # Wait longer to see if it prints
                    logger.info(f"Check your printer - did it print from {test_port}?")
                else:
                    logger.error(f"Failed to connect to {test_port}")
            except Exception as e:
                logger.error(f"Error testing {test_port}: {e}")
            finally:
                if test_printer.is_connected():
                    test_printer.disconnect()
                time.sleep(1)
        
        logger.info("\nTest complete. Check which port printed and use that one.")
        exit(0)
    
    try:
        if not printer.connect():
            logger.error("Failed to connect to printer")
            exit(1)
        
        if args.raw_hex:
            # Send raw hex bytes
            try:
                hex_bytes = bytes.fromhex(args.raw_hex.replace(' ', '').replace('0x', ''))
                logger.info(f"Sending raw hex: {args.raw_hex}")
                printer.send_raw(hex_bytes, debug=args.debug)
            except ValueError as e:
                logger.error(f"Invalid hex string: {e}")
                exit(1)
        elif args.test:
            logger.info("Sending test print...")
            if args.simple:
                printer.print_simple("Test Print\nBluetooth COM Bridge\nWorking!", encoding=args.encoding)
            else:
                # Use TSPL by default for label printers
                printer.print_label(
                    "Test Print\nBluetooth COM Bridge\nWorking!",
                    use_tspl=not args.no_tspl,
                    label_width_mm=args.label_width,
                    label_height_mm=args.label_height
                )
        elif args.text:
            logger.info(f"Printing text: {args.text}")
            if args.simple:
                printer.print_simple(args.text, encoding=args.encoding)
            else:
                # Use TSPL by default for label printers
                printer.print_label(
                    args.text,
                    use_tspl=not args.no_tspl,
                    label_width_mm=args.label_width,
                    label_height_mm=args.label_height
                )
        elif args.pdf:
            logger.info(f"Printing PDF: {args.pdf}")
            printer.print_pdf(
                args.pdf,
                label_width_mm=args.label_width,
                label_height_mm=args.label_height,
                dpi=args.dpi,
                auto_crop=not args.no_crop,
                auto_rotate=not args.no_rotate,
                invert=not args.no_invert
            )
        elif args.file:
            logger.info(f"Printing file: {args.file}")
            printer.send_file(args.file)
        else:
            logger.warning("No print command specified. Use --text, --file, or --test")
            parser.print_help()
    
    except KeyboardInterrupt:
        logger.info("Interrupted by user")
    finally:
        printer.disconnect()

