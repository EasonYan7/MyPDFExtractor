# MyPDFExtractor
Here's a comprehensive **README.md** file for your GitHub repository:

```markdown
# MyPDFExtractor

**MyPDFExtractor** is a powerful PyQt6-based desktop application designed to help users visually select and extract text from PDF files. The selected content can be exported to Excel for further analysis. The tool supports batch processing, drag-and-drop functionality, and synchronization across multiple PDFs.

---

## 🚀 Features

- **Drag-and-Drop Support**  
  Easily load multiple PDF files by dragging them into the application window.
  
- **Text Selection with Visual Feedback**  
  Select areas within PDFs using a mouse-drawn rectangle and view extracted content instantly.

- **Batch Processing**  
  Synchronize selections across multiple PDFs for consistent extraction.

- **Excel Export**  
  Save extracted content with structured columns for easy data analysis.

- **User-Friendly Navigation**  
  Quickly browse through PDF pages with dedicated buttons.

- **Selection Management**  
  View, delete, and clear selections with an intuitive interface.

---

## 📥 Installation

### Prerequisites

Ensure you have the following installed:

- **Python 3.8+**
- **Dependencies** (install using `pip`):

  ```bash
  pip install -r requirements.txt
  ```

### Run the Application

1. Clone the repository:

   ```bash
   git clone https://github.com/EasonYan7/pdf-extractor.git
   cd pdf-extractor
   ```

2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. Launch the application:

   ```bash
   python pdf_read.py
   ```

---

## 🏗️ Build an Executable (Windows)

If you want to create a standalone `.exe`:

1. Install PyInstaller:

   ```bash
   pip install pyinstaller
   ```

2. Build the executable:

   ```bash
   pyinstaller --onefile --windowed --name PDFExtractor pdf_read.py
   ```

3. The `.exe` file will be located in the `dist` folder.

---

## 🖥️ Usage Instructions

### Load PDF Files

- **Drag-and-Drop PDFs**: Simply drag files into the left panel.  
- **Open PDF Folder**: Click `Open PDF Folder` to load multiple PDFs at once.

### Navigate Through PDFs

- Use **Previous Page** and **Next Page** buttons to move between pages.

### Select Text Areas

1. Click and drag to draw a rectangle over the text in the PDF preview.
2. The extracted content will appear in the **Extraction Preview** section.
3. Selected areas are listed in the table, with options to delete.

### Sync Selections Across PDFs

- Click **Sync to All PDFs** to apply the same selections to all loaded PDFs.

### Export Selections

- Click **Export Selections** to save extracted text as an Excel file.

---

## 📂 Project Structure

```
pdf-extractor/
│-- pdf_read.py             # Main application script
│-- README.md               # Project documentation
│-- requirements.txt        # Dependencies list
│-- assets/                 # Any image/icons
│-- dist/                    # Compiled executable output (after build)
```

---

## ⚠️ Known Issues & Limitations

1. **OCR Requirement**  
   - The application extracts selectable text only. Scanned PDFs need OCR processing before selection.

2. **Performance on Large PDFs**  
   - For very large files, the rendering might take longer.

3. **Coordinate Scaling**  
   - Selections are based on the displayed scale; ensure consistency across PDFs.

---

## 🛠️ Technologies Used

- **[PyQt6](https://www.riverbankcomputing.com/software/pyqt/)** - For the GUI framework  
- **[PyMuPDF](https://pymupdf.readthedocs.io/)** - To handle PDF text extraction  
- **[Pandas](https://pandas.pydata.org/)** - To format and export data to Excel  

---

## 📖 License

This project is licensed under the Apache License Version 2.0, January 2004

---

## 👨‍💻 Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository.
2. Create a new feature branch:  
   ```bash
   git checkout -b feature-name
   ```
3. Commit changes:  
   ```bash
   git commit -m "Add new feature"
   ```
4. Push to your fork and create a pull request.

---

## 📝 Contact

For any inquiries or suggestions, feel free to contact:

- **Author:** Eason
- **Email:** mozziebolt@gmail.com
- **GitHub:** EasonYan7

---

Enjoy using **MyPDFExtractor**! 😊
