
# 🧙‍♂️ FileWitch

**FileWitch** is a versatile Python command-line tool that empowers you to convert and merge files between various formats effortlessly. Whether you're dealing with CSV, Excel, Word, PDF, or plain text files, FileWitch has got you covered.

---

![Downloads](https://static.pepy.tech/personalized-badge/filewitch?period=total\&units=international_system\&left_color=grey\&right_color=blue)

---

## ✨ Features

* 🔄 **Bidirectional Conversion**: Seamlessly convert between:

  * CSV ↔️ Excel (`.csv` ↔️ `.xlsx`)
  * Text ↔️ Word (`.txt` ↔️ `.docx`)
  * Text → PDF (`.txt` → `.pdf`)
  * Word → PDF (`.docx` → `.pdf`)
  * PowerPoint → Word (`.pptx` → `.docx`)
  * PowerPoint → PDF (`.pptx` → `.pdf`)
* 📁 **File Duplication**: Copy files within the same format
* ➕ **File Merging**: Merge multiple files of the same format without losing content or overlapping data
* 🛠️ **Simple CLI Interface**: Perform conversions and merges with straightforward commands
* 🖼️ **Image Support**: Preserve images during conversions
* 📝 **Format Preservation**: Maintain formatting in DOCX and PDF conversions
* 📊 **Complex Elements**: Handle tables, lists, and other complex elements

## 🚀 Installation

Ensure you have Python 3.7 or higher installed. Then, install FileWitch using pip:

```bash
pip install filewitch
```

## 🧪 Usage

### 🔧 Command-Line Interface

Perform file conversions directly from your terminal:

```bash
# Convert CSV to Excel
filewitch convert data.csv xlsx

# Merge multiple CSV files into one
filewitch merge file1.csv file2.csv file3.csv merged.csv

# Merge multiple Word files into one
filewitch merge doc1.docx doc2.docx merged.docx
```

### 🧠 Python API

Integrate FileWitch into your Python projects:

```python
from filewitch import (
    csv_to_xlsx,
    xlsx_to_csv,
    txt_to_docx,
    docx_to_txt,
    txt_to_pdf,
    docx_to_pdf,
    pptx_to_docx,
    pptx_to_pdf,
    merge_files
)

# Merge CSV files
merge_files(['data1.csv', 'data2.csv'], 'merged.csv')

# Merge Word documents
merge_files(['chapter1.docx', 'chapter2.docx'], 'book.docx')
```

## 📂 Supported Conversions & Merges

| Format  | Conversion Support | Merge Support |
| ------- | ------------------ | ------------- |
| `.csv`  | ✅                  | ✅             |
| `.xlsx` | ✅                  | ✅             |
| `.txt`  | ✅                  | ✅             |
| `.docx` | ✅                  | ✅             |
| `.pdf`  | ✅                  | ✅             |
| `.pptx` | ✅                  | ✅             |

## ⚙️ Dependencies

FileWitch leverages the following Python libraries:

* `pandas`: For CSV and Excel operations
* `openpyxl`: For Excel file manipulation
* `python-docx`: For Word document handling
* `python-pptx`: For PowerPoint file handling
* `reportlab`: For PDF generation
* `click`: For command-line interface
* `Pillow`: For image processing
* `docx2pdf`: For high-quality DOCX to PDF conversion
* `pptx2pdf`: For high-quality PPTX to PDF conversion
* `PyPDF2`: For merging PDF files


## 📄 License

This project is licensed under the [MIT License](LICENSE).

## 👨‍💻 Author

* **Raktim Kalita** - [GitHub](https://github.com/Rktim)
* Email: [raktimkalita.ai@gmail.com](mailto:raktimkalita.ai@gmail.com)

