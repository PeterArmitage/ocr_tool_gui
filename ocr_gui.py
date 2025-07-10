import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import os
import sys
from pathlib import Path
import subprocess
import platform
import tempfile
import io
from datetime import datetime

print("Starting Enhanced OCR GUI...")

class EnhancedOCRGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("OCR Text Extractor - Enhanced with Export & Clipboard")
        self.root.geometry("900x700")
        
        # Check tesseract availability with detailed diagnostics
        self.tesseract_available, self.tesseract_info = self.check_tesseract_detailed()
        
        self.setup_ui()
        self.setup_clipboard_monitoring()
        
    def check_tesseract_detailed(self):
        """Enhanced tesseract checking with detailed diagnostics"""
        info = []
        
        # Check if pytesseract is installed
        try:
            import pytesseract
            info.append("✓ pytesseract module imported successfully")
        except ImportError as e:
            info.append(f"✗ pytesseract not found: {e}")
            return False, info
        
        # Check if PIL is available
        try:
            from PIL import Image
            info.append("✓ PIL (Pillow) imported successfully")
        except ImportError as e:
            info.append(f"✗ PIL/Pillow not found: {e}")
            return False, info
        
        # Check tesseract executable
        try:
            version = pytesseract.get_tesseract_version()
            info.append(f"✓ Tesseract version: {version}")
        except Exception as e:
            info.append(f"✗ Tesseract executable not found: {e}")
            
            # Try to find tesseract executable
            possible_paths = self.find_tesseract_paths()
            if possible_paths:
                info.append("Possible Tesseract locations found:")
                for path in possible_paths:
                    info.append(f"  - {path}")
                    
                # Try setting the path to the first found location
                try:
                    pytesseract.pytesseract.tesseract_cmd = possible_paths[0]
                    version = pytesseract.get_tesseract_version()
                    info.append(f"✓ Tesseract working with path: {possible_paths[0]}")
                    info.append(f"✓ Version: {version}")
                    return True, info
                except:
                    info.append("✗ Found paths don't work")
            else:
                info.append("No tesseract executable found in common locations")
            
            return False, info
        
        return True, info
    
    def find_tesseract_paths(self):
        """Find possible tesseract executable locations"""
        possible_paths = []
        
        if platform.system() == "Windows":
            # Common Windows installation paths
            windows_paths = [
                r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                r"C:\Users\%USERNAME%\AppData\Local\Tesseract-OCR\tesseract.exe",
            ]
            
            for path in windows_paths:
                expanded_path = os.path.expandvars(path)
                if os.path.exists(expanded_path):
                    possible_paths.append(expanded_path)
        
        elif platform.system() == "Darwin":  # macOS
            mac_paths = [
                "/usr/local/bin/tesseract",
                "/opt/homebrew/bin/tesseract",
                "/usr/bin/tesseract",
            ]
            
            for path in mac_paths:
                if os.path.exists(path):
                    possible_paths.append(path)
        
        else:  # Linux
            linux_paths = [
                "/usr/bin/tesseract",
                "/usr/local/bin/tesseract",
                "/opt/tesseract/bin/tesseract",
            ]
            
            for path in linux_paths:
                if os.path.exists(path):
                    possible_paths.append(path)
        
        # Try to find using 'which' command
        try:
            result = subprocess.run(['which', 'tesseract'], 
                                  capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                path = result.stdout.strip()
                if path not in possible_paths:
                    possible_paths.append(path)
        except:
            pass
        
        return possible_paths
    
    def setup_clipboard_monitoring(self):
        """Setup clipboard monitoring for image paste"""
        # Bind Ctrl+V to paste function
        self.root.bind('<Control-v>', self.paste_from_clipboard)
        self.root.bind('<Control-V>', self.paste_from_clipboard)
        
        # Make the window focusable to receive keyboard events
        self.root.focus_set()
    
    def paste_from_clipboard(self, event=None):
        """Handle clipboard paste for images"""
        try:
            from PIL import Image, ImageGrab
            
            # Try to get image from clipboard
            clipboard_image = ImageGrab.grabclipboard()
            
            if clipboard_image is not None:
                # Save temporary image
                temp_path = os.path.join(tempfile.gettempdir(), 
                                       f"clipboard_image_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
                clipboard_image.save(temp_path)
                
                # Set the file path and process
                self.file_var.set(temp_path)
                self.status_label.config(text="Clipboard image loaded! Processing...", fg="blue")
                self.root.update()
                
                # Process the image
                self.process_file()
                
                # Clean up temp file after processing
                try:
                    os.remove(temp_path)
                except:
                    pass
                
                messagebox.showinfo("Success", "Clipboard image processed successfully!")
            else:
                messagebox.showwarning("No Image", "No image found in clipboard. Copy an image first (Print Screen, Snipping Tool, etc.)")
        
        except ImportError:
            messagebox.showerror("Error", "PIL/Pillow not installed. Cannot access clipboard.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to paste from clipboard: {str(e)}")
    
    def setup_ui(self):
        # Title
        title = tk.Label(self.root, text="OCR Text Extractor - Enhanced with Export & Clipboard", 
                        font=("Arial", 16, "bold"))
        title.pack(pady=10)
        
        # Top button frame
        top_btn_frame = tk.Frame(self.root)
        top_btn_frame.pack(pady=5)
        
        # Diagnostics button
        diag_btn = tk.Button(top_btn_frame, text="Show Diagnostics", 
                           command=self.show_diagnostics,
                           bg="orange", fg="white")
        diag_btn.pack(side="left", padx=5)
        
        # Clipboard paste button
        clipboard_btn = tk.Button(top_btn_frame, text="📋 Paste from Clipboard (Ctrl+V)", 
                                command=self.paste_from_clipboard,
                                bg="purple", fg="white")
        clipboard_btn.pack(side="left", padx=5)
        
        # File selection frame
        file_frame = tk.Frame(self.root)
        file_frame.pack(pady=10, padx=20, fill="x")
        
        tk.Label(file_frame, text="File:").pack(side="left")
        
        self.file_var = tk.StringVar()
        self.file_entry = tk.Entry(file_frame, textvariable=self.file_var, width=50)
        self.file_entry.pack(side="left", padx=5)
        
        browse_btn = tk.Button(file_frame, text="Browse", command=self.browse_file)
        browse_btn.pack(side="left", padx=5)
        
        # OCR Configuration frame
        config_frame = tk.LabelFrame(self.root, text="OCR Configuration")
        config_frame.pack(pady=10, padx=20, fill="x")
        
        # Language selection
        lang_frame = tk.Frame(config_frame)
        lang_frame.pack(pady=5, fill="x")
        
        tk.Label(lang_frame, text="Language:").pack(side="left")
        self.lang_var = tk.StringVar(value="eng")
        lang_combo = tk.OptionMenu(lang_frame, self.lang_var, "eng", "por", "spa", "fra", "deu")
        lang_combo.pack(side="left", padx=5)
        
        # PSM (Page Segmentation Mode) selection
        psm_frame = tk.Frame(config_frame)
        psm_frame.pack(pady=5, fill="x")
        
        tk.Label(psm_frame, text="Page Segmentation Mode:").pack(side="left")
        self.psm_var = tk.StringVar(value="6")
        psm_combo = tk.OptionMenu(psm_frame, self.psm_var, "3", "6", "7", "8", "13")
        psm_combo.pack(side="left", padx=5)
        
        # Process button
        self.process_btn = tk.Button(self.root, text="Extract Text", 
                                   command=self.process_file,
                                   bg="green", fg="white", font=("Arial", 12))
        self.process_btn.pack(pady=10)
        
        # Results area
        tk.Label(self.root, text="Extracted Text:", font=("Arial", 12, "bold")).pack(pady=(20,5))
        
        self.results_text = scrolledtext.ScrolledText(self.root, width=80, height=12)
        self.results_text.pack(pady=10, padx=20, fill="both", expand=True)
        
        # Enhanced Export Frame
        export_frame = tk.LabelFrame(self.root, text="Export Options")
        export_frame.pack(pady=10, padx=20, fill="x")
        
        # Export format selection
        format_frame = tk.Frame(export_frame)
        format_frame.pack(pady=5, fill="x")
        
        tk.Label(format_frame, text="Export Format:").pack(side="left")
        self.export_format = tk.StringVar(value="txt")
        format_combo = ttk.Combobox(format_frame, textvariable=self.export_format, width=15)
        format_combo['values'] = ('txt', 'pdf', 'docx', 'rtf', 'html')
        format_combo.pack(side="left", padx=5)
        
        # Export buttons
        export_btn_frame = tk.Frame(export_frame)
        export_btn_frame.pack(pady=5)
        
        self.export_btn = tk.Button(export_btn_frame, text="📁 Export As...", 
                                   command=self.export_results,
                                   bg="blue", fg="white")
        self.export_btn.pack(side="left", padx=5)
        
        quick_save_btn = tk.Button(export_btn_frame, text="💾 Quick Save (.txt)", 
                                 command=self.quick_save_txt,
                                 bg="darkblue", fg="white")
        quick_save_btn.pack(side="left", padx=5)
        
        clear_btn = tk.Button(export_btn_frame, text="🗑️ Clear Results", 
                            command=self.clear_results,
                            bg="red", fg="white")
        clear_btn.pack(side="left", padx=5)
        
        # Status
        status_text = "Ready - Try Ctrl+V to paste screenshots!" if self.tesseract_available else "ERROR: Tesseract not configured properly"
        self.status_label = tk.Label(self.root, text=status_text, 
                                   fg="green" if self.tesseract_available else "red")
        self.status_label.pack(pady=5)
        
        # Instructions
        instructions = tk.Label(self.root, 
                              text="💡 Tip: Use Print Screen or Snipping Tool, then press Ctrl+V to paste and process images instantly!",
                              font=("Arial", 9), fg="gray")
        instructions.pack(pady=2)
        
        if not self.tesseract_available:
            self.process_btn.config(state="disabled")
    
    def show_diagnostics(self):
        """Show detailed diagnostics information"""
        diag_window = tk.Toplevel(self.root)
        diag_window.title("Tesseract Diagnostics")
        diag_window.geometry("600x400")
        
        text_widget = scrolledtext.ScrolledText(diag_window, wrap=tk.WORD)
        text_widget.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Add system info
        text_widget.insert(tk.END, "=== SYSTEM INFORMATION ===\n")
        text_widget.insert(tk.END, f"Platform: {platform.system()} {platform.release()}\n")
        text_widget.insert(tk.END, f"Python: {sys.version}\n\n")
        
        # Add tesseract diagnostics
        text_widget.insert(tk.END, "=== TESSERACT DIAGNOSTICS ===\n")
        for info in self.tesseract_info:
            text_widget.insert(tk.END, f"{info}\n")
        
        # Check additional libraries
        text_widget.insert(tk.END, "\n=== EXPORT CAPABILITIES ===\n")
        
        # Check export libraries
        libs_to_check = [
            ("python-docx", "Word document export"),
            ("reportlab", "PDF export"),
            ("fpdf", "Alternative PDF export"),
            ("pypandoc", "Universal document conversion")
        ]
        
        for lib, purpose in libs_to_check:
            try:
                __import__(lib.replace('-', '_'))
                text_widget.insert(tk.END, f"✓ {lib} - {purpose}\n")
            except ImportError:
                text_widget.insert(tk.END, f"✗ {lib} - {purpose} (install with: pip install {lib})\n")
        
        text_widget.insert(tk.END, "\n=== INSTALLATION INSTRUCTIONS ===\n")
        text_widget.insert(tk.END, "Core OCR packages:\n")
        text_widget.insert(tk.END, "pip install pytesseract pillow\n\n")
        text_widget.insert(tk.END, "Export packages (optional):\n")
        text_widget.insert(tk.END, "pip install python-docx reportlab fpdf pypandoc\n\n")
        
        if platform.system() == "Windows":
            text_widget.insert(tk.END, "Windows - Tesseract Installation:\n")
            text_widget.insert(tk.END, "1. Download from: https://github.com/tesseract-ocr/tesseract/releases\n")
            text_widget.insert(tk.END, "2. Install the .exe file\n")
            text_widget.insert(tk.END, "3. Add to PATH or set pytesseract.pytesseract.tesseract_cmd\n")
        elif platform.system() == "Darwin":
            text_widget.insert(tk.END, "macOS - Tesseract Installation:\n")
            text_widget.insert(tk.END, "1. Install Homebrew if needed\n")
            text_widget.insert(tk.END, "2. Run: brew install tesseract\n")
        else:
            text_widget.insert(tk.END, "Linux - Tesseract Installation:\n")
            text_widget.insert(tk.END, "1. Run: sudo apt-get install tesseract-ocr\n")
        
        text_widget.config(state="disabled")
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select file",
            filetypes=[
                ("All supported", "*.pdf *.png *.jpg *.jpeg *.tiff *.bmp *.gif"),
                ("PDF files", "*.pdf"),
                ("Images", "*.png *.jpg *.jpeg *.tiff *.bmp *.gif"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.file_var.set(file_path)
    
    def process_file(self):
        file_path = self.file_var.get()
        if not file_path:
            messagebox.showerror("Error", "Please select a file first or paste from clipboard")
            return
        
        if not os.path.exists(file_path):
            messagebox.showerror("Error", "File not found")
            return
        
        self.status_label.config(text="Processing...", fg="orange")
        self.process_btn.config(state="disabled")
        self.root.update()
        
        try:
            if file_path.lower().endswith('.pdf'):
                result = self.process_pdf(file_path)
            else:
                result = self.process_image(file_path)
            
            self.results_text.delete(1.0, tk.END)
            self.results_text.insert(tk.END, result)
            self.status_label.config(text="Complete! Ready to export.", fg="green")
            
        except Exception as e:
            error_msg = f"Processing failed: {str(e)}"
            messagebox.showerror("Error", error_msg)
            self.results_text.delete(1.0, tk.END)
            self.results_text.insert(tk.END, f"ERROR: {error_msg}\n\nClick 'Show Diagnostics' for troubleshooting help.")
            self.status_label.config(text="Error", fg="red")
        finally:
            self.process_btn.config(state="normal")
    
    def process_image(self, image_path):
        try:
            import pytesseract
            from PIL import Image
            
            # Load and preprocess image
            img = Image.open(image_path)
            
            # Convert to RGB if necessary
            if img.mode != 'RGB':
                img = img.convert('RGB')
            
            # Configure tesseract options
            custom_config = f'--oem 3 --psm {self.psm_var.get()} -l {self.lang_var.get()}'
            
            # Extract text
            text = pytesseract.image_to_string(img, config=custom_config)
            
            result = f"=== IMAGE OCR RESULTS ===\n"
            result += f"File: {os.path.basename(image_path)}\n"
            result += f"Image size: {img.size}\n"
            result += f"Language: {self.lang_var.get()}\n"
            result += f"PSM: {self.psm_var.get()}\n"
            result += f"Characters found: {len(text)}\n"
            result += f"Processed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            result += "="*50 + "\n\n"
            
            if text.strip():
                result += text
            else:
                result += "No text found in image. Try:\n"
                result += "- Different Page Segmentation Mode\n"
                result += "- Different language setting\n"
                result += "- Image preprocessing (better contrast, resolution)\n"
            
            return result
            
        except Exception as e:
            return f"Error processing image: {str(e)}\n\nTroubleshooting:\n1. Check if file is a valid image\n2. Verify Tesseract installation\n3. Check file permissions"
    
    def process_pdf(self, pdf_path):
        try:
            # Check if PyMuPDF is available
            try:
                import fitz  # PyMuPDF
            except ImportError:
                return "Error: PyMuPDF not installed. Install with: pip install PyMuPDF"
            
            import pytesseract
            from PIL import Image
            import io
            
            doc = fitz.open(pdf_path)
            all_text = []
            
            all_text.append(f"=== PDF OCR RESULTS ===")
            all_text.append(f"File: {os.path.basename(pdf_path)}")
            all_text.append(f"Pages: {len(doc)}")
            all_text.append(f"Language: {self.lang_var.get()}")
            all_text.append(f"PSM: {self.psm_var.get()}")
            all_text.append(f"Processed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            all_text.append("="*50)
            
            for page_num in range(len(doc)):
                page = doc[page_num]
                
                # Get regular text
                regular_text = page.get_text()
                if regular_text.strip():
                    all_text.append(f"\n=== Page {page_num + 1} - Regular Text ===")
                    all_text.append(regular_text)
                
                # Get images and OCR them
                images = page.get_images()
                if images:
                    all_text.append(f"\n=== Page {page_num + 1} - Images ({len(images)} found) ===")
                    
                    for img_index, img in enumerate(images):
                        try:
                            xref = img[0]
                            pix = fitz.Pixmap(doc, xref)
                            
                            if pix.n - pix.alpha < 4:
                                img_data = pix.tobytes("ppm")
                            else:
                                pix = fitz.Pixmap(fitz.csRGB, pix)
                                img_data = pix.tobytes("ppm")
                            
                            pil_img = Image.open(io.BytesIO(img_data))
                            
                            # Configure tesseract
                            custom_config = f'--oem 3 --psm {self.psm_var.get()} -l {self.lang_var.get()}'
                            ocr_text = pytesseract.image_to_string(pil_img, config=custom_config)
                            
                            if ocr_text.strip():
                                all_text.append(f"\n--- Image {img_index + 1} ---")
                                all_text.append(ocr_text)
                            else:
                                all_text.append(f"\n--- Image {img_index + 1} (no text found) ---")
                            
                            pix = None
                        except Exception as e:
                            all_text.append(f"\n--- Image {img_index + 1} Error ---")
                            all_text.append(str(e))
            
            doc.close()
            return "\n".join(all_text) if all_text else "No text found in PDF"
            
        except Exception as e:
            return f"Error processing PDF: {str(e)}\n\nTroubleshooting:\n1. Install PyMuPDF: pip install PyMuPDF\n2. Check if PDF is readable\n3. Verify Tesseract installation"
    
    def quick_save_txt(self):
        """Quick save as text file with timestamp"""
        text = self.results_text.get(1.0, tk.END)
        if not text.strip():
            messagebox.showwarning("Warning", "No text to save")
            return
        
        # Generate filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"ocr_results_{timestamp}.txt"
        
        file_path = filedialog.asksaveasfilename(
            title="Quick Save as Text",
            initialname=filename,
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(text)
                messagebox.showinfo("Success", f"Saved to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save: {str(e)}")
    
    def export_results(self):
        """Export results in selected format"""
        text = self.results_text.get(1.0, tk.END)
        if not text.strip():
            messagebox.showwarning("Warning", "No text to export")
            return
        
        format_type = self.export_format.get()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Define file type mappings
        file_types = {
            'txt': ("Text files", "*.txt"),
            'pdf': ("PDF files", "*.pdf"),
            'docx': ("Word documents", "*.docx"),
            'rtf': ("Rich Text Format", "*.rtf"),
            'html': ("HTML files", "*.html")
        }
        
        file_path = filedialog.asksaveasfilename(
            title=f"Export as {format_type.upper()}",
            initialname=f"ocr_results_{timestamp}.{format_type}",
            defaultextension=f".{format_type}",
            filetypes=[file_types[format_type], ("All files", "*.*")]
        )
        
        if file_path:
            try:
                if format_type == 'txt':
                    self.export_txt(file_path, text)
                elif format_type == 'pdf':
                    self.export_pdf(file_path, text)
                elif format_type == 'docx':
                    self.export_docx(file_path, text)
                elif format_type == 'rtf':
                    self.export_rtf(file_path, text)
                elif format_type == 'html':
                    self.export_html(file_path, text)
                
                messagebox.showinfo("Success", f"Exported to {file_path}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export: {str(e)}")
    
    def export_txt(self, file_path, text):
        """Export as plain text"""
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(text)
    
    def export_pdf(self, file_path, text):
        """Export as PDF"""
        try:
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import letter
            from reportlab.lib.utils import simpleSplit
            
            c = canvas.Canvas(file_path, pagesize=letter)
            width, height = letter
            
            # Set up text formatting
            y = height - 50
            line_height = 12
            margin = 50
            
            # Split text into lines
            lines = text.split('\n')
            
            for line in lines:
                if y < margin:
                    c.showPage()
                    y = height - 50
                
                # Handle long lines
                if len(line) > 80:
                    wrapped_lines = simpleSplit(line, "Helvetica", 10, width - 2 * margin)
                    for wrapped_line in wrapped_lines:
                        c.drawString(margin, y, wrapped_line)
                        y -= line_height
                else:
                    c.drawString(margin, y, line)
                    y -= line_height
            
            c.save()
            
        except ImportError:
            # Fallback to fpdf if reportlab not available
            try:
                from fpdf import FPDF
                
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=10)
                
                # Add text line by line
                for line in text.split('\n'):
                    pdf.cell(0, 5, line.encode('latin-1', 'replace').decode('latin-1'), ln=True)
                
                pdf.output(file_path)
                
            except ImportError:
                raise ImportError("PDF export requires 'reportlab' or 'fpdf' library. Install with: pip install reportlab")
    
    def export_docx(self, file_path, text):
        """Export as Word document"""
        try:
            from docx import Document
            
            doc = Document()
            doc.add_heading('OCR Results', 0)
            doc.add_paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            doc.add_paragraph(text)
            doc.save(file_path)
            
        except ImportError:
            raise ImportError("Word export requires 'python-docx' library. Install with: pip install python-docx")
    
    def export_rtf(self, file_path, text):
        """Export as Rich Text Format"""
        # Simple RTF format
        rtf_header = r"{\rtf1\ansi\deff0 {\fonttbl {\f0 Times New Roman;}}}"
        rtf_content = text.replace('\n', '\\par ')
        rtf_footer = "}"
        
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(f"{rtf_header}\\f0\\fs24 {rtf_content}{rtf_footer}")
    
    def export_html(self, file_path, text):
        """Export as HTML"""
        html_content = f"""<!DOCTYPE html>
<html>
<head>
    <title>OCR Results</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        .header {{ color: #333; border-bottom: 2px solid #ddd; padding-bottom: 10px; }}
        .content {{ white-space: pre-wrap; margin-top: 20px; }}
        .timestamp {{ color: #666; font-size: 0.9em; }}
    </style>
</head>
<body>
    <h1 class="header">OCR Results</h1>
    <p class="timestamp">Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
    <div class="content">{text.replace('<', '&lt;').replace('>', '&gt;')}</div>
</body>
</html>"""
        
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
    
    def clear_results(self):
        self.results_text.delete(1.0, tk.END)
        self.status_label.config(text="Ready - Try Ctrl+V to paste screenshots!", fg="green")
    
    def run(self):
        print("Starting Enhanced GUI main loop...")
        print("Features available:")
        print("- File OCR (PDF, Images)")
        print("- Clipboard paste (Ctrl+V)")
        print("- Multiple export formats (TXT, PDF, DOCX, RTF, HTML)")
        self.root.mainloop()
        print("GUI closed.")

if __name__ == "__main__":
    try:
        app = EnhancedOCRGUI()
        app.run()
    except Exception as e:
        print(f"Error starting GUI: {e}")
        input("Press Enter to exit...")