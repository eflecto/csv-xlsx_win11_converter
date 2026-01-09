# 📊 CSV to XLSX Converter

A beautiful, feature-rich Windows 11 utility for converting CSV files to Excel (XLSX) format with extensive customization options.

![Python Version](https://img.shields.io/badge/python-3.9%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Platform](https://img.shields.io/badge/platform-Windows%2011-lightgrey)

## ✨ Features

### 📁 File Management
- **Batch Processing**: Convert multiple CSV files at once
- **Folder Import**: Add all CSV files from a folder
- **Drag & Drop**: Easy file management (coming soon)

### ⚙️ CSV Settings
- **Multiple Encodings**: UTF-8, UTF-8-BOM, Latin-1, CP1251, CP1252, ISO-8859-1, ASCII
- **Custom Delimiters**: Comma, Semicolon, Tab, Pipe
- **Skip Rows**: Ignore header/metadata rows
- **Custom Sheet Names**: Name your Excel sheets

### 🎨 Excel Styling
- **Auto-fit Columns**: Automatically adjust column widths
- **Freeze Header**: Keep headers visible while scrolling
- **Auto-Filters**: Add Excel filtering capability
- **Styled Headers**: Custom header background and text colors
- **Zebra Stripes**: Alternating row colors for readability
- **Color Presets**: Quick apply preset color schemes
- **Cell Borders**: Clean, professional borders

### 📊 Preview
- **Live Preview**: Preview CSV data before conversion
- **First 100 Rows**: Quick data inspection

### 🖥️ User Interface
- **Modern Dark/Light Theme**: Toggle between themes
- **Progress Tracking**: Real-time conversion progress
- **Multi-threaded**: Non-blocking UI during conversion

## 🚀 Installation

### Prerequisites
- Python 3.9 or higher
- Windows 10/11 (optimized for Windows 11)

### Quick Start

1. **Clone the repository**


2. **Create a virtual environment (recommended)**
```bash
python -m venv venv
venv\Scripts\activate
```

3. **Install dependencies**
```bash
pip install -r requirements.txt
```

4. **Run the application**
```bash
python main.py
```

## 📦 Building Executable

Create a standalone `.exe` file for distribution:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=icon.ico --name="CSV-to-XLSX-Converter" main.py
```

The executable will be in the `dist` folder.

## 🎯 Usage

1. **Add Files**: Click "Add Files" or "Add Folder" to select CSV files
2. **Configure Settings**: Adjust CSV parsing and Excel styling options
3. **Preview** (Optional): Check the data preview tab
4. **Select Output**: Choose output folder for XLSX files
5. **Convert**: Click the green "Convert to XLSX" button

## 📸 Screenshots

### Main Interface
```
┌────────────────────────────────────────────────────────────┐
│  📊 CSV to XLSX Converter                    [Dark Mode ●] │
├────────────────────────────────────────────────────────────┤
│  [📁 Files] [⚙️ CSV Settings] [🎨 Excel Styling] [📊 Preview] │
├────────────────────────────────────────────────────────────┤
│                                                            │
│  [➕ Add Files] [📁 Add Folder] [🗑️ Remove] [🧹 Clear]      │
│                                                            │
│  ┌──────────────────────────────────────────────────────┐ │
│  │ 1. C:\Data\sales_2024.csv                            │ │
│  │ 2. C:\Data\inventory.csv                             │ │
│  │ 3. C:\Data\customers.csv                             │ │
│  └──────────────────────────────────────────────────────┘ │
│                                                            │
│  📂 Output: [C:\Users\...\Documents        ] [Browse]      │
│                                                            │
├────────────────────────────────────────────────────────────┤
│  ████████████████████░░░░░░░░░░░░░░░░░░░░░░  45%           │
│                    Converting file 2 of 3...               │
│  ┌────────────────────────────────────────────────────────┐│
│  │               🚀 Convert to XLSX                       ││
│  └────────────────────────────────────────────────────────┘│
└────────────────────────────────────────────────────────────┘
```

## 🔧 Configuration Options

### Encoding Options
| Encoding | Description | Use Case |
|----------|-------------|----------|
| `utf-8` | Standard Unicode | Default, most files |
| `utf-8-sig` | UTF-8 with BOM | Excel-exported CSVs |
| `cp1251` | Cyrillic | Russian Windows files |
| `cp1252` | Western European | Windows files |
| `latin-1` | ISO Latin-1 | Legacy files |

### Delimiter Options
| Delimiter | Symbol | Common Usage |
|-----------|--------|--------------|
| Comma | `,` | Standard CSV |
| Semicolon | `;` | European locale |
| Tab | `\t` | TSV files |
| Pipe | `\|` | Data exports |

## 🏗️ Project Structure

```
csv-to-xlsx-converter/
├── main.py              # Main application
├── requirements.txt     # Python dependencies
├── README.md           # This file
├── LICENSE             # MIT License
├── .gitignore          # Git ignore rules
├── CHANGELOG.md        # Version history
└── assets/             # Icons and images (optional)
    └── icon.ico
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📋 Roadmap

- [ ] Drag and drop file support
- [ ] Custom color picker dialog
- [ ] Save/load configuration presets
- [ ] Command-line interface (CLI) mode
- [ ] Multiple sheets per file
- [ ] Data type detection and formatting
- [ ] Excel formulas support
- [ ] Localization (multi-language)

## ❓ FAQ

**Q: Why are my Cyrillic characters showing incorrectly?**
A: Try using `cp1251` or `utf-8-sig` encoding.

**Q: Can I convert multiple files with different settings?**
A: Currently, all files use the same settings. Batch with different settings is on the roadmap.

**Q: The application freezes during conversion?**
A: This shouldn't happen as conversion runs in a separate thread. Please report if it occurs.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) - Modern UI framework
- [Pandas](https://pandas.pydata.org/) - Data manipulation
- [OpenPyXL](https://openpyxl.readthedocs.io/) - Excel file handling

## 📧 Contact

Your Name - [@yourtwitter](https://twitter.com/yourtwitter) - email@example.com

Project Link: [https://github.com/yourusername/csv-to-xlsx-converter](https://github.com/yourusername/csv-to-xlsx-converter)

---

⭐ Star this repo if you find it useful!
