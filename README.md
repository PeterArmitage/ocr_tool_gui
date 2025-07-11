# OCR Text Extractor - Enhanced GUI

# This is a work in progress, there are some bugs!

A powerful, user-friendly OCR (Optical Character Recognition) tool with clipboard integration and multiple export formats. Extract text from images, PDFs, and screenshots with ease!

## ✨ Features

- **Multi-format OCR**: Extract text from images (PNG, JPG, TIFF, BMP, GIF) and PDF files
- **Clipboard Integration**: Paste screenshots directly with Ctrl+V for instant OCR processing
- **Multiple Export Formats**: Save results as TXT, PDF, DOCX, RTF, or HTML
- **Language Support**: English, Portuguese, Spanish, French, German, and more
- **Page Segmentation Options**: Fine-tune OCR accuracy with different PSM modes
- **User-friendly GUI**: Clean, intuitive interface with drag-and-drop functionality
- **Diagnostics Tool**: Built-in troubleshooting and installation verification

## 🚀 Quick Start

### 1. Clone the Repository

```bash
git clone https://github.com/PeterArmitage/ocr-gui-enhanced.git
cd ocr-gui-enhanced
```

### 2. Install Dependencies

```bash
# Core dependencies (required)
pip install pytesseract pillow PyMuPDF

# Export libraries (optional, for advanced export formats)
pip install python-docx reportlab fpdf

# Or just install the requirements file
pip install -r requirements.txt
```

### 3. Install Tesseract OCR Engine

#### Windows

1. Download Tesseract from [GitHub Releases](https://github.com/tesseract-ocr/tesseract/releases)
2. Install the `.exe` file
3. Add Tesseract to your system PATH, or the tool will help you locate it

#### macOS

```bash
# Install Homebrew if you haven't already
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# Install Tesseract
brew install tesseract
```

#### Linux (Ubuntu/Debian)

```bash
sudo apt-get update
sudo apt-get install tesseract-ocr
```

### 4. Run the Application

```bash
python ocr_gui.py
```

## 🖥️ Usage

### Basic OCR Process

1. **Launch the application**
2. **Select a file** using the "Browse" button or drag-and-drop
3. **Configure settings** (language, page segmentation mode)
4. **Click "Extract Text"** to process the file
5. **Export results** in your preferred format

### Clipboard OCR (Super Fast!)

1. **Take a screenshot** (Print Screen, Snipping Tool, etc.)
2. **Press Ctrl+V** in the OCR tool or click the purple clipboard button
3. **Text is automatically extracted** and displayed!

### Export Options

- **Quick Save (.txt)**: Instant text file with timestamp
- **Export As...**: Choose from multiple formats:
  - **TXT**: Plain text file
  - **PDF**: Professional formatted document
  - **DOCX**: Microsoft Word document
  - **RTF**: Rich Text Format (universal)
  - **HTML**: Web-friendly format with styling

## ⚙️ Configuration Options

### Languages Supported

- **English** (eng) - Default
- **Portuguese** (por)
- **Spanish** (spa)
- **French** (fra)
- **German** (deu)

### Page Segmentation Modes (PSM)

- **PSM 3**: Fully automatic page segmentation
- **PSM 6**: Uniform block of text (default)
- **PSM 7**: Single text line
- **PSM 8**: Single word
- **PSM 13**: Raw line (no assumptions)

## 📁 File Formats Supported

### Input

- **Images**: PNG, JPG, JPEG, TIFF, BMP, GIF
- **Documents**: PDF (with embedded images)
- **Clipboard**: Any image copied to clipboard

### Output

- **TXT**: Plain text
- **PDF**: Formatted document
- **DOCX**: Microsoft Word
- **RTF**: Rich Text Format
- **HTML**: Web format

## 🔧 Dependencies

### Required

```
pytesseract>=0.3.10
Pillow>=9.0.0
PyMuPDF>=1.20.0
```

### Optional (for advanced exports)

```
python-docx>=0.8.11
reportlab>=3.6.0
fpdf>=3.0.0
```

## 🛠️ Troubleshooting

### Common Issues

#### "Tesseract not found"

- **Windows**: Use the built-in diagnostics tool to locate Tesseract
- **macOS/Linux**: Ensure Tesseract is installed and in PATH
- **All platforms**: Run diagnostics from the application for detailed help

#### "No text found in image"

- Try different Page Segmentation Mode (PSM) settings
- Ensure image has good contrast and resolution
- Check if the correct language is selected

#### Export errors

- Check if optional libraries are installed for specific formats
- Use the diagnostics tool to see which export formats are available

### Getting Help

1. **Click "Show Diagnostics"** in the application for detailed system information
2. **Check the console output** for error messages
3. **Verify all dependencies** are installed correctly

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request. Areas for improvement:

- Additional language support
- Image preprocessing options
- Batch processing capabilities
- OCR confidence scoring
- Cloud storage integration

## 🙏 Acknowledgments

- **Tesseract OCR** - Google's OCR engine
- **PyTesseract** - Python wrapper for Tesseract
- **Pillow** - Python Imaging Library
- **PyMuPDF** - PDF processing library

## 📊 System Requirements

- **Python**: 3.7 or higher
- **Operating System**: Windows, macOS, or Linux
- **RAM**: 2GB minimum (4GB recommended for large files)
- **Storage**: 50MB for application + dependencies

## 🔄 Version History

- **v2.0**: Added clipboard integration and multiple export formats
- **v1.0**: Basic OCR functionality with GUI

---

**Made with ❤️ for easy text extraction**

_Star this repository if you find it useful!_
