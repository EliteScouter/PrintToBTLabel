# 🏷️ Bluetooth Label Printer

A modern Windows desktop app to print shipping labels from PDF files to cheap Bluetooth thermal label printers (MVGGES, TSC, Xprinter, etc.).

**No drivers needed** - connects directly via Bluetooth COM port using TSPL protocol.

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

## ✨ Features

- 📁 **PDF to Label** - Print shipping labels from PDF files (eBay, Amazon, USPS, etc.)
- 👁️ **Live Preview** - See exactly what will print before sending
- ✂️ **Smart Crop** - Auto-detects label boundaries, or draw manual crop area
- 🔄 **Auto-Rotate** - Automatically rotates image to fit label orientation
- 📐 **Preset Sizes** - 4×6", 4×4", 2×1" or custom dimensions
- 🔌 **Easy Connection** - Select COM port from dropdown, auto-detects available ports
- 🎨 **Invert Colors** - Toggle color inversion if your prints come out reversed
- 🌙 **Dark Mode UI** - Modern, easy on the eyes

## 🖨️ Supported Printers

Works with most cheap Bluetooth thermal label printers that use **TSPL (TSC Printer Language)**:

- MVGGES label printers
- TSC label printers
- Xprinter thermal printers
- Most Chinese Bluetooth label printers from Amazon/AliExpress

## 📥 Installation

### Option 1: Download Executable (Recommended)

Download `LabelPrinter.exe` from the [Releases](../../releases) page - no Python required!

### Option 2: Run from Source

1. Install Python 3.8 or higher
2. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/PrintToBTLabel.git
   cd PrintToBTLabel
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Run the app:
   ```bash
   python label_printer_app.py
   ```

## 🚀 Quick Start

1. **Pair your printer** via Windows Bluetooth settings
2. **Launch the app** (LabelPrinter.exe or `python label_printer_app.py`)
3. **Select COM port** - Usually COM5 or COM6 for Bluetooth
4. **Load a PDF** - Click "Select PDF" and choose your shipping label
5. **Print!** - Click the big green Print button

## 🛠️ Building the Executable

To create your own standalone .exe:

```bash
python build_exe.py
```

The executable will be created at `dist/LabelPrinter.exe`

## ⚙️ Settings

| Setting | Description |
|---------|-------------|
| **Label Size** | Choose preset (4×6, 4×4, 2×1) or enter custom mm dimensions |
| **Auto-crop** | Automatically detect and crop to label boundaries |
| **Auto-rotate** | Rotate image to match label orientation |
| **Flip vertical** | Flip the image upside down |
| **Invert colors** | Invert black/white (use if print colors are reversed) |
| **Manual crop** | Draw a rectangle to select exactly what to print |

## 🔧 Troubleshooting

### Printer not showing up?
- Make sure it's paired in Windows Bluetooth settings
- Check if a COM port was assigned (Device Manager → Ports)
- Click the refresh button (↻) next to port dropdown

### Nothing prints?
- Verify the COM port is correct (try both if you see two)
- Check printer has paper/labels loaded
- Make sure printer is powered on and in range

### Print is inverted/negative?
- Toggle the "Invert colors" checkbox

### Print is upside down?
- Toggle the "Flip vertical" checkbox

### Wrong area printing?
- Enable "Manual crop" and draw a selection around the area you want

## 💻 Command Line Usage

For automation or scripting, use the CLI:

```bash
# List available COM ports
python bt_printer.py --list-ports

# Print a PDF
python bt_printer.py --pdf label.pdf --port COM5

# Print with custom label size
python bt_printer.py --pdf label.pdf --label-width 101 --label-height 152

# Print text directly
python bt_printer.py --text "Hello World" --port COM5
```

## 🐍 Using as a Python Module

```python
from bt_printer import BluetoothPrinter

# Print a PDF label
with BluetoothPrinter(port="COM5") as printer:
    printer.print_pdf("shipping_label.pdf")

# Print text
with BluetoothPrinter(port="COM5") as printer:
    printer.print_label("Order #12345\nShip to: John Doe")
```

## 📋 Requirements

- Windows 10/11
- Python 3.8+ (if running from source)
- Bluetooth adapter
- Compatible thermal label printer

## 🤝 Contributing

Contributions welcome! Feel free to:
- Report bugs
- Suggest features
- Submit pull requests

## 📄 License

MIT License - feel free to use this in your own projects!

## 🙏 Acknowledgments

- Built with [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) for the modern UI
- PDF rendering via [PyMuPDF](https://github.com/pymupdf/PyMuPDF)
