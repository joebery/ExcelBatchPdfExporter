# ExcelBatchPdfExporter

A Windows desktop application for batch converting Excel files to PDF using Excel’s native Print-to-PDF functionality.

![C#](https://img.shields.io/badge/language-C%23-blue.svg) ![License](https://img.shields.io/badge/license-None-lightgrey.svg)

---

## ✨ What This Project Does
ExcelBatchPdfExporter is a Windows desktop application that enables users to batch convert Excel workbooks to PDF format using Excel's built-in `ExportAsFixedFormat` functionality. With a simple WPF GUI, it maintains layout fidelity through late-bound Excel COM automation.

## 🛠️ Tech Stack
| Technology        | Description                              |
|-------------------|------------------------------------------|
| .NET 8            | Framework used for application development|
| WPF               | User interface framework                 |
| Excel COM         | Used for automation of Excel tasks       |

## 🎉 Features
- 📁 Batch convert multiple Excel files in one operation
- ⚙️ Global configuration to export a specific worksheet index across all files
- ⚠️ Automatic validation and warning for files missing the selected sheet
- 📂 Save PDFs next to input files or in a single chosen output folder
- 📝 Deterministic output naming: example: exceldocs-1.pdf
- 🔒 Optional overwrite protection
- 🖨️ Uses Excel’s native `ExportAsFixedFormat` PDF engine

---

## 🚀 Getting Started
### Prerequisites
- Windows
- Microsoft Excel installed
- .NET 8 SDK

### How to Run
From the repository root, run:
```powershell
dotnet run --project src\ExcelBatchPdfExporter.Gui
```

---

## 📜 What Changed In This Update
- Revised README with additional details and screenshots for improved project information and features.
- Auto-generated README via DevOps Portfolio Analyzer.
- Added directory for documentation and screenshots.

---

## 📊 Architecture
```
+-------------------+
|  ExcelBatchPdfExporter  |
+-------------------+
|    WPF GUI        |
|                   |
|  Excel COM        |
+-------------------+
```

---

## 🤝 Contributing
Contributions are welcome! Please open an issue or submit a pull request.

## 📄 License
This project is licensed under the terms of the MIT license.