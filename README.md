# 📄 docx-file-splitter

[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](LICENSE)
[![Python: 3.8+](https://img.shields.io/badge/Python-3.8%2B-brightgreen.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)]()
[![Privacy](https://img.shields.io/badge/Privacy-100%25%20Local%20%26%20Offline-success.svg)]()
[![Dependencies](https://img.shields.io/badge/Dependencies-python--docx-orange.svg)]()

A fast, lightweight, and **lossless Microsoft Word (`.docx`) document splitter**. It divides any Word document into **$N$ equal segments based on Memory Distribution** while preserving **100% of formatting, tables, images, styles, and layout**.

Works entirely **offline on your machine** with zero cloud uploads. Includes both an **interactive wizard** and a **sleek Dark-Mode Desktop GUI**.

---

## ⚡ Quick Start in 30 Seconds

### 1. Install Dependencies
Only requires `python-docx` (Python 3.8+ required):
```bash
pip install -r requirements.txt
```

### 2. Run the Splitter
```bash
python main.py
```

### 3. Answer 2 Simple Prompts:
```text
📂 Step 1: Select Word Document (.docx)
   -> Type path, drag & drop your file, or press Enter to open File Browser

✂️  Step 2: Number of Segments (Memory Distribution)
   -> Enter how many parts you need (e.g. 2, 3, 5)

✨ Done! Files are saved in a clean folder next to your original document.
```

> **Windows 1-Click Launchers**:
> - Double-click **`run_cli.bat`** to launch the interactive terminal prompt.
> - Double-click **`run_gui.bat`** to launch the Desktop GUI window.

---

## 🧠 Why Split by Memory Distribution?

Standard splitters only count paragraphs, which causes major issues when documents contain **screenshots, photos, diagrams, or large tables**:
- If Chapter 1 has 5 large images and Chapter 2 has none, a paragraph split might produce one **70 MB** file and one **200 KB** file!
- **Word Document Splitter solves this with Memory Distribution**:
  1. It analyzes embedded media payloads, image sizes, and content weights.
  2. It calculates optimal cut-points to ensure each split segment has an **evenly balanced file size in megabytes (MB)**.
  3. It cleans up unused image relationships in each segment to keep output files compact and lightweight.

Perfect for keeping documents under **email attachment limits (e.g. 25 MB)**, complying with upload limits for AI systems and web portals, or sharing large reports.

---

## 🎮 How to Use

### Option 1: Interactive Wizard (Easiest)
Simply run:
```bash
python main.py
```
- Supports **Drag & Drop** directly into your terminal.
- Press **Enter** on the file prompt to launch the native file picker dialog.
- Live progress bar showing each segment as it is created.
- Displays a clean breakdown table showing word counts and file sizes.
- Offers to automatically open the destination folder in File Explorer / Finder.

### Option 2: Modern Desktop GUI
If you prefer a visual dark-mode window:
```bash
python main.py --gui
```
*(or double-click `run_gui.bat` on Windows)*

- Visual document statistics (Total Words, Paragraphs, Tables, File Size).
- Quick preset buttons (`2`, `3`, `4`, `5`, `10` segments).
- Real-time progress bar.
- 1-click button to open the output folder.

### Option 3: Automated Command-Line (For Scripts)
Pass the document path and segment count directly:
```bash
# Split report.docx into 3 segments based on memory distribution
python main.py "report.docx" 3

# Specify custom output folder
python main.py -f "thesis.docx" -n 4 -o "./my_segments"
```

#### CLI Options Reference
```text
python main.py [file] [segments] [options]

Arguments:
  file                       Path to input .docx document
  segments                   Number of segments to split into

Options:
  -h, --help                 Show help message and exit
  -g, --gui                  Launch the Desktop GUI window
  -f, --file-path PATH       Path to the input .docx file
  -n, --num-segments N       Number of segments to create
  -m, --method METHOD        Strategy: 'memory' (default), 'words', 'blocks', 'headings'
  -o, --output DIR           Custom output directory for split files
  -a, --advanced             Interactive mode with advanced algorithm and folder prompts
```

---

## 🔒 Privacy & Security Guarantee

- **100% Offline & Private**: All document parsing and OpenXML manipulation occurs locally in memory on your computer. No document content, text, or metadata is ever transmitted over the network.
- **Repository Safe**: Pre-configured with strict `.gitignore` rules that automatically prevent any personal Word documents (`*.docx`, `*.doc`), PDFs, and generated split folders (`*_split/`) from being staged or committed to GitHub.

---

## 🐍 Python API Usage

You can also import and use the core engine directly in your own Python projects:

```python
from word_splitter_core import WordSplitter

# 1. Analyze document statistics
stats = WordSplitter.analyze_document("sample.docx")
print(f"File Size: {stats['file_size']:,} bytes | Blocks: {stats['total_blocks']}")

# 2. Split into 3 balanced segments by memory
result = WordSplitter.split_document(
    file_path="sample.docx",
    num_parts=3,
    method="memory",  # 'memory', 'words', 'blocks', or 'headings'
    output_dir="./split_output"
)

# 3. Inspect generated segments
for part in result["parts"]:
    print(f"Segment {part['part_number']}: {part['file_name']} ({part['words']} words, {part['file_size']} bytes)")
```

---

## 📁 Repository Structure

```
Word file splitter/
├── main.py                     # Primary entry point (interactive CLI or GUI)
├── word_splitter_core.py       # Core lossless OpenXML engine & memory algorithms
├── word_splitter_cli.py        # Streamlined 2-step interactive terminal interface
├── word_splitter_gui.py        # Modern Dark-Mode GUI (Tkinter)
├── splitter.py                 # Unified runner with flag dispatching
│
├── run_cli.bat                 # Windows 1-click CLI launcher
├── run_gui.bat                 # Windows 1-click GUI launcher
├── run_cli.sh                  # macOS/Linux 1-click CLI launcher
├── run_gui.sh                  # macOS/Linux 1-click GUI launcher
│
├── create_sample_document.py   # Generates sample Word documents for testing
├── test_splitter.py            # Automated unit test suite
├── requirements.txt            # Python dependencies (python-docx)
├── pyproject.toml              # Modern package metadata
├── LICENSE                     # MIT License
├── CONTRIBUTING.md             # Contribution guidelines
└── .gitignore                  # Security rules (protects personal documents)
```

---

## 🧪 Testing

### Run Automated Unit Tests
```bash
python test_splitter.py
```

### Generate a Test Document
To generate a realistic sample Word document with chapters, tables, and styled text without using any personal files:
```bash
python create_sample_document.py
```
This generates `example_document.docx` (automatically ignored by git so it won't be committed).

---

## 📄 License

This project is licensed under the [GNU General Public License v3.0 (GPL-3.0)](LICENSE).
