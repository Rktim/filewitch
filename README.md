# 🧙‍♂️ FileWitch

**FileWitch** is a versatile Python command-line tool that empowers you to convert files between various formats effortlessly. Whether you're dealing with CSV, Excel, Word, PDF, or plain text files, FileWitch has got you covered.

---
![PyPI - Downloads](https://img.shields.io/pypi/dm/filewitch)

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
* 🛠️ **Simple CLI Interface**: Perform conversions with straightforward commands
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

# Convert Excel to CSV
filewitch convert data.xlsx csv

# Convert Text to Word
filewitch convert notes.txt docx

# Convert Word to Text
filewitch convert report.docx txt

# Convert Text to PDF
filewitch convert notes.txt pdf

# Convert Word to PDF
filewitch convert report.docx pdf

# Convert PowerPoint to Word
filewitch convert presentation.pptx docx

# Convert PowerPoint to PDF
filewitch convert presentation.pptx pdf

# Copy a file within the same format
filewitch convert document.txt txt
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
    pptx_to_pdf
)

# Convert CSV to Excel
csv_to_xlsx('input.csv', 'output.xlsx')

# Convert Excel to CSV
xlsx_to_csv('input.xlsx', 'output.csv')

# Convert Text to Word
txt_to_docx('input.txt', 'output.docx')

# Convert Word to Text
docx_to_txt('input.docx', 'output.txt')

# Convert Text to PDF
txt_to_pdf('input.txt', 'output.pdf')

# Convert Word to PDF
docx_to_pdf('input.docx', 'output.pdf')

# Convert PowerPoint to Word
pptx_to_docx('input.pptx', 'output.docx')

# Convert PowerPoint to PDF
pptx_to_pdf('input.pptx', 'output.pdf')
```

## 📂 Supported Conversions

| Source Format | Target Format | Description          | Features Preserved |
| ------------- | ------------- | -------------------- | ------------------ |
| `.csv`        | `.xlsx`       | CSV to Excel         | Data formatting    |
| `.xlsx`       | `.csv`        | Excel to CSV         | Data structure     |
| `.txt`        | `.docx`       | Text to Word         | Headings, formatting|
| `.docx`       | `.txt`        | Word to Text         | Text content, image placeholders |
| `.txt`        | `.pdf`        | Text to PDF          | Text formatting    |
| `.docx`       | `.pdf`        | Word to PDF          | Formatting, images |
| `.pptx`       | `.docx`       | PowerPoint to Word   | Text, images, formatting |
| `.pptx`       | `.pdf`        | PowerPoint to PDF    | Slides, images, formatting |
| `.txt`        | `.txt`        | Copy Text file       | All content        |
| `.csv`        | `.csv`        | Copy CSV file        | All data           |
| `.xlsx`       | `.xlsx`       | Copy Excel file      | All data and formatting |
| `.docx`       | `.docx`       | Copy Word file       | All content and formatting |
| `.pdf`        | `.pdf`        | Copy PDF file        | All content        |
| `.pptx`       | `.pptx`       | Copy PowerPoint file | All content        |

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

## 🤝 Contributing

Contributions are welcome! To contribute:

1. Fork the repository
2. Create a new branch: `git checkout -b feature-name`
3. Make your changes and commit them: `git commit -m 'Add new feature'`
4. Push to the branch: `git push origin feature-name`
5. Open a pull request

## 📄 License

This project is licensed under the [MIT License](LICENSE).

## 👨‍💻 Author

- **Raktim Kalita** - [GitHub](https://github.com/Rktim)
- Email: raktimkalita.ai@gmail.com 
