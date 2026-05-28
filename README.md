## **Image → JPEG Converter**

A full-featured desktop application for batch converting images of any format to JPEG. Built with Python and Tkinter — works on Windows, macOS, and Linux.

---
<img width="777" height="882" alt="Social_preview" src="https://github.com/user-attachments/assets/6f29a61a-65ec-49a5-811a-a9f912d16230" />
---
 
## ✨ Features
 
- **Batch conversion** — Add unlimited images and convert them all in one click
- **Import from documents** — Extract embedded images directly from PDF, DOCX, PPTX, XLSX, and ZIP/CBZ files
- **Drag & drop** — Drop image files, document files, or entire folders onto the app
- **Sequential naming** — Outputs named `Image_1.jpg`, `Image_2.jpg`... in the exact order you arrange them
- **Custom prefix** — Choose your own naming prefix (e.g. `Photo_1.jpg`, `Scan_1.jpg`)
- **Reorder control** — Move files up or down before converting to control the output sequence
- **Thumbnail preview** — See a small preview of each image in the file list
- **Quality control** — JPEG quality slider from 10 to 100
- **Threaded conversion** — UI stays fully responsive during large batch jobs
- **Dark / Light mode** — Sage & Slate theme with full dark mode support
- **Duplicate support** — Same image can appear multiple times; each gets its own output file
- **Auto-installs dependencies** — No manual pip installs needed on first run
---
 
## 📋 Supported Formats
 
| Type | Formats |
|------|---------|
| **Images** | PNG, JPG, BMP, GIF, TIFF, WEBP, ICO, PPM, TGA, HEIC, HEIF |
| **Documents** | PDF, DOCX, PPTX, XLSX, ZIP, CBZ |
 
---
  
## ⚙️ Requirements
 
- **Python 3.6 or higher**
  - Download: https://www.python.org/downloads
  - ⚠️ Windows: check **"Add Python to PATH"** during install
  - ⚠️ macOS: install via Homebrew for full Unicode/emoji support (see below)
The following libraries are **auto-installed on first run** — no manual setup needed:
 
| Library | Purpose |
|---------|---------|
| Pillow | Image processing and conversion |
| PyMuPDF | Extract images from PDF files |
| python-docx | Extract images from Word documents |
| python-pptx | Extract images from PowerPoint files |
| openpyxl | Extract images from Excel files |
| tkinterdnd2 | Drag and drop support |
 
---
 
## 🚀 Getting Started
 
**Clone the repository**
```bash
git clone https://github.com/rjkjigneshparmar/image-to-jpeg-converter.git
cd image-to-jpeg-converter
```
 
**Run the app**
 
```bash
# Windows
py convert_to_jpeg.py
 
# macOS / Linux
python3 convert_to_jpeg.py
```
 
Dependencies install automatically on first run.
 
---
 
## 🍎 macOS Setup
 
macOS ships with an outdated Tcl/Tk that causes a Unicode error. Fix it by installing Python via Homebrew:
 
```bash
# Install Homebrew (if not already installed)
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
 
# Install Python with modern Tcl/Tk
brew install python-tk
 
# Verify Tcl/Tk version (must be 8.6+)
python3 -c "import tkinter; print(tkinter.TclVersion)"
 
# Run the app
python3 convert_to_jpeg.py
```
 
---
 
## 📖 How to Use
 
1. **Add files** — Click `+ Add Images` to pick images, `Import File` to extract from documents, or drag & drop files/folders directly onto the list
2. **Reorder** — Use `Up` / `Down` buttons to set the output order (first item = `Image_1.jpg`)
3. **Set output folder** — Click `Browse…` and choose where JPEGs will be saved
4. **Set prefix** *(optional)* — Change `Image` to any name you prefer
5. **Adjust quality** *(optional)* — Drag the quality slider (default: 90)
6. **Convert** — Click `Convert to JPEG` and watch the progress log
7. **Open results** — Click `Open Output` to jump straight to your files
---
 
## 🏗 Build as Standalone App
 
To distribute without requiring Python:
 
**Windows — builds `ImageToJPEG.exe`**
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "ImageToJPEG" convert_to_jpeg.py
# Output: dist/ImageToJPEG.exe
```
 
**macOS — builds `ImageToJPEG.app`**
```bash
pip3 install pyinstaller
pyinstaller --onefile --windowed --name "ImageToJPEG" convert_to_jpeg.py
# Output: dist/ImageToJPEG.app
```
 
> ⚠️ The app must be built on the same OS it will run on. You cannot build a Mac app on Windows or vice versa.
 
---
 
## 📁 Project Structure
 
```
image-to-jpeg-converter/
├── convert_to_jpeg.py   # Main application script
├── HOW_TO_USE.txt       # Full usage guide
├── HOW_TO_BUILD_APP.txt # Build instructions for Windows & macOS
├── README.md            # This file
└── screenshots/         # App screenshots (optional)
```
 
---
 
## 🛠 Troubleshooting
 
| Problem | Fix |
|---------|-----|
| `Python was not found` on Windows | Use `py convert_to_jpeg.py` or disable the Microsoft Store alias in Settings → Apps → Advanced app settings → App execution aliases |
| `TclError: character above U+FFFF` on macOS | Install Python via Homebrew: `brew install python-tk` |
| `ModuleNotFoundError` | Run `pip install Pillow pymupdf python-docx python-pptx openpyxl tkinterdnd2` |
| No images found after Import File | The document may contain only vector graphics (not raster images) |
| App crashes silently after build | Remove `--windowed` from PyInstaller command to see error output |
 
---
 
## 📄 License
 
This project is open source. Feel free to use, modify, and distribute.
 
---
 
## 👤 Author
 
**Jignesh Parmar**
GitHub: [@rjkjigneshparmar](https://github.com/rjkjigneshparmar)
 
