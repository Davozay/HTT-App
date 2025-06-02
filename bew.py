import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import os
import subprocess
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, LabelFrame


# ==================== BACKEND ====================
class OCRProcessor:
    """Handles image processing and text extraction"""

    def __init__(self, tesseract_path=None):
        # Try multiple common Tesseract installation paths
        possible_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            r"C:\Users\{}\AppData\Local\Programs\Tesseract-OCR\tesseract.exe".format(os.getenv('USERNAME')),
            "tesseract"  # If it's in PATH
        ]

        if tesseract_path:
            possible_paths.insert(0, tesseract_path)

        # Test each path
        tesseract_found = False
        for path in possible_paths:
            try:
                if path == "tesseract":
                    # Test if tesseract is in PATH
                    import subprocess
                    result = subprocess.run(['tesseract', '--version'],
                                            capture_output=True, text=True, timeout=10)
                    if result.returncode == 0:
                        pytesseract.pytesseract.tesseract_cmd = 'tesseract'
                        tesseract_found = True
                        print(f"DEBUG: Using Tesseract from PATH")
                        break
                else:
                    if os.path.exists(path):
                        pytesseract.pytesseract.tesseract_cmd = path
                        tesseract_found = True
                        print(f"DEBUG: Found Tesseract at: {path}")
                        break
            except Exception as e:
                print(f"DEBUG: Failed to test path {path}: {e}")
                continue

        if not tesseract_found:
            raise ValueError(
                "Tesseract not found! Please install Tesseract-OCR from:\n"
                "https://github.com/UB-Mannheim/tesseract/wiki\n"
                f"Tried paths: {possible_paths}"
            )

        self.config = '--oem 3 --psm 6'

    @staticmethod
    def preprocess_image(image_path):
        """Enhances image for better OCR results"""
        try:
            print(f"DEBUG: preprocess_image called with: {repr(image_path)}")

            # Add validation for image_path
            if image_path is None:
                raise ValueError("Image path is None")

            if not os.path.exists(image_path):
                raise ValueError(f"Image file does not exist: {image_path}")

            print("DEBUG: About to open image")
            img = Image.open(image_path)
            print(f"DEBUG: Opened image: {img}")

            print("DEBUG: Converting to grayscale")
            img = img.convert('L')  # Grayscale

            print("DEBUG: Enhancing contrast")
            enhancer = ImageEnhance.Contrast(img)
            img = enhancer.enhance(2)  # Increase contrast

            print("DEBUG: Applying sharpen filter")
            # Suppress the type warning by using type: ignore comment
            img = img.filter(ImageFilter.SHARPEN)  # type: ignore

            print(f"DEBUG: Final processed image: {img}")
            return img
        except Exception as e:
            print(f"DEBUG: Exception in preprocess_image: {e}")
            raise ValueError(f"Image processing failed: {str(e)}")

    def extract_text(self, image_path, preprocess=True):
        """Extracts text from image using OCR"""
        try:
            # Debug print statements
            print(f"DEBUG: image_path = {repr(image_path)}")
            print(f"DEBUG: image_path type = {type(image_path)}")
            print(f"DEBUG: preprocess = {preprocess}")

            # Add validation for image_path
            if image_path is None:
                raise ValueError("Image path is None")
            if image_path == "":
                raise ValueError("Image path is empty string")
            if not isinstance(image_path, (str, bytes, os.PathLike)):
                raise ValueError(f"Image path is wrong type: {type(image_path)}")

            if preprocess:
                print("DEBUG: About to preprocess image")
                img = OCRProcessor.preprocess_image(image_path)  # Call as static method
                print(f"DEBUG: Preprocessed image = {repr(img)}")
                # Additional check to ensure img is not None
                if img is None:
                    raise ValueError("Preprocessed image is None")
            else:
                print("DEBUG: Opening image without preprocessing")
                if not os.path.exists(image_path):
                    raise ValueError(f"Image file does not exist: {image_path}")
                img = Image.open(image_path)
                print(f"DEBUG: Opened image = {repr(img)}")

            # Ensure img is valid before passing to pytesseract
            if img is None:
                raise ValueError("Image object is None")

            print("DEBUG: About to call pytesseract.image_to_string")
            print(f"DEBUG: Tesseract command: {pytesseract.pytesseract.tesseract_cmd}")

            # Test tesseract command before using it
            try:
                if isinstance(pytesseract.pytesseract.tesseract_cmd, str):
                    if os.path.exists(
                            pytesseract.pytesseract.tesseract_cmd) or pytesseract.pytesseract.tesseract_cmd == 'tesseract':
                        # Test tesseract
                        test_result = subprocess.run([pytesseract.pytesseract.tesseract_cmd, '--version'],
                                                     capture_output=True, text=True, timeout=5)
                        if test_result.returncode != 0:
                            raise ValueError(f"Tesseract test failed: {test_result.stderr}")
                    else:
                        raise ValueError(f"Tesseract executable not found at: {pytesseract.pytesseract.tesseract_cmd}")
                else:
                    raise ValueError(
                        f"Tesseract command is not a string: {type(pytesseract.pytesseract.tesseract_cmd)}")
            except subprocess.TimeoutExpired:
                raise ValueError("Tesseract test timed out")
            except Exception as e:
                raise ValueError(f"Tesseract validation failed: {e}")

            text = pytesseract.image_to_string(img, config=self.config)
            print(f"DEBUG: Extracted text length = {len(text)}")
            return text.strip()
        except Exception as e:
            print(f"DEBUG: Exception in extract_text: {e}")
            raise ValueError(f"Text extraction failed: {str(e)}")


class HandwritingToTextBackend:
    """Main backend controller"""

    def __init__(self, tesseract_path=None):
        self.ocr = OCRProcessor(tesseract_path)

    def process_image(self, image_path):
        """Handles complete image-to-text processing"""
        # Add validation for image_path
        if not image_path:
            return False, "No image path provided"

        if not os.path.exists(image_path):
            return False, "File not found"

        if not image_path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
            return False, "Unsupported file format"

        try:
            extracted_text = self.ocr.extract_text(image_path)
            return True, extracted_text
        except Exception as e:
            return False, str(e)

    @staticmethod
    def save_text(text, output_path):
        """Saves text to file"""
        try:
            # Add validation for parameters
            if not text:
                return False, "No text to save"
            if not output_path:
                return False, "No output path provided"

            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(text)
            return True, "Text saved successfully"
        except Exception as e:
            return False, f"Error saving file: {str(e)}"


# ==================== FRONTEND ====================
class TextExtractorApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Handwriting to Text Converter")
        self.master.state('zoomed')
        self.master.configure(bg="#000000")

        # Initialize backend
        self.backend = HandwritingToTextBackend()

        # Styling
        self.button_style = {"bg": "#6C757D", "fg": "white", "font": ("Arial", 12), "width": 20}
        self.text_style = {"bg": "white", "font": ("Arial", 12)}
        self.red_button_style = {"bg": "#DC3545", "fg": "white", "font": ("Arial", 12), "width": 20}

        # Store current image path
        self.current_image = None

        # Initialize GUI attributes (to fix the warning)
        self.upload_label = None
        self.text_display = None

        # GUI Components
        self.create_widgets()

    def create_widgets(self):
        """Create all interface components"""
        # Header
        header = tk.Label(self.master, text="Handwriting to Text Converter",
                          font=("Arial", 18, "bold"), bg="#000000", fg="white")
        header.pack(pady=10)

        # Image Upload Section
        upload_frame = LabelFrame(self.master, text="Image Upload",
                                  font=("Arial", 12, "bold"), bg="#1C1C1C", fg="white",
                                  padx=10, pady=10)
        upload_frame.pack(pady=10, padx=20, fill=tk.X)

        self.upload_label = tk.Label(upload_frame,
                                     text="🖼️ No image selected\nSupported formats: PNG, JPG, JPEG, BMP",
                                     font=("Arial", 12), bg="#1C1C1C", fg="white")
        self.upload_label.pack(pady=5)

        upload_btn = tk.Button(upload_frame, text="📁 Select Image",
                               command=self.upload_image, **self.button_style)
        upload_btn.pack(pady=5)

        # Extract Text Button
        extract_btn = tk.Button(self.master, text="🔄 Convert to Text",
                                command=self.extract_text, **self.button_style)
        extract_btn.pack(pady=10)

        # Extracted Text Section
        text_frame = LabelFrame(self.master, text="Extracted Text",
                                font=("Arial", 12, "bold"), bg="#1C1C1C", fg="white",
                                padx=10, pady=10)
        text_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=20)

        self.text_display = scrolledtext.ScrolledText(text_frame, width=80,
                                                      height=10, **self.text_style)
        self.text_display.pack(pady=5, fill=tk.BOTH, expand=True)

        # Save Button
        save_btn = tk.Button(self.master, text="💾 Save Text",
                             command=self.save_text, **self.button_style)
        save_btn.pack(pady=10)

        # Clear Button
        clear_btn = tk.Button(self.master, text="Clear Text",
                              command=self.clear_text, **self.red_button_style)
        clear_btn.pack(pady=5)

    def upload_image(self):
        """Handle image selection"""
        file_path = filedialog.askopenfilename(
            title="Select an Image",
            filetypes=[
                ("Image Files", "*.png *.jpg *.jpeg *.bmp"),
                ("PNG files", "*.png"),
                ("JPEG files", "*.jpg *.jpeg"),
                ("BMP files", "*.bmp"),
                ("All files", "*.*")
            ]
        )

        print(f"DEBUG: Selected file path: {repr(file_path)}")

        if file_path:
            self.current_image = file_path
            self.upload_label.config(
                text=f"🖼️ Selected: {os.path.basename(file_path)}"
            )
            print(f"DEBUG: Set self.current_image to: {repr(self.current_image)}")
        else:
            # Handle case where user cancels file selection
            self.current_image = None
            print("DEBUG: User cancelled file selection")

    def extract_text(self):
        """Handle text extraction from image"""
        if not self.current_image:
            messagebox.showerror("Error", "Please select an image first!")
            return

        try:
            # Show processing message
            self.master.config(cursor="wait")
            self.master.update()

            success, result = self.backend.process_image(self.current_image)

            if success:
                self.text_display.delete('1.0', tk.END)
                self.text_display.insert(tk.END, result)
                messagebox.showinfo("Success", "Text extracted successfully!")
            else:
                messagebox.showerror("Error", f"Text extraction failed: {result}")

        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error: {str(e)}")
        finally:
            # Reset cursor
            self.master.config(cursor="")

    def save_text(self):
        """Save extracted text to file"""
        text_content = self.text_display.get('1.0', tk.END).strip()

        if not text_content:
            messagebox.showerror("Error", "No text to save!")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text file", "*.txt")]
        )

        if file_path:
            success, message = self.backend.save_text(text_content, file_path)
            if success:
                messagebox.showinfo("Success", message)
            else:
                messagebox.showerror("Error", message)

    def clear_text(self):
        """Clear the text display"""
        self.text_display.delete('1.0', tk.END)


# ==================== MAIN ====================
if __name__ == "__main__":
    root = tk.Tk()
    app = TextExtractorApp(root)
    root.mainloop()