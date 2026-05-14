# ExcelBatchPdfExporter

A Windows desktop application for batch converting Excel files to PDF using Excel’s native Print-to-PDF functionality.

![Language](https://img.shields.io/badge/language-C%23-blue.svg)

## 🎯 What This Project Does
ExcelBatchPdfExporter allows users to batch convert multiple Excel workbooks into PDF files with a simple WPF GUI. It leverages Excel's built-in `ExportAsFixedFormat` functionality while ensuring layout fidelity through late-bound Excel COM automation.

## ⚙️ Tech Stack
| Technology      | Description                      |
|-----------------|----------------------------------|
| C#              | Programming language used        |
| .NET 8         | Framework for building applications|
| WPF             | UI framework for desktop apps    |

## ✨ Features
- 📄 Batch convert multiple Excel files in one operation
- ⚙️ Global configuration to export a specific worksheet index across all files
- ⚠️ Automatic validation and warning for files missing the selected sheet
- 📁 Save PDFs next to input files or in a single chosen output folder
- 🔄 Deterministic output naming (e.g., `exceldocs-1.pdf`)
- 🔒 Optional overwrite protection
- 📥 Uses Excel’s native `ExportAsFixedFormat` PDF engine

## 🚀 Getting Started
### Prerequisites
- Windows
- Microsoft Excel installed
- .NET 8 SDK

### Commands
To run the application, execute:
```powershell
dotnet run --project src\ExcelBatchPdfExporter.Gui
```

## 🛠️ What Changed In This Update
Recent commits include:
- Added a directory for documentation and screenshots.
- Updated the README file to provide more information about the project.
- Initial commit of the Excel batch PDF exporter application.

## 🏗️ Architecture
```
+----------------------+   +-------------------------+
|   ExcelBatchPdfExporter  |<--|   Microsoft Excel      |
|   (WPF GUI)          |   |   (COM Automation)     |
+----------------------+   +-------------------------+
           |                            |
           |                            |
           |                            |
+----------------------+   +-------------------------+
|      PDF Files       |   |  User Input Files       |
+----------------------+   +-------------------------+
```

## 🤝 Contributing
Contributions are welcome! Please submit a pull request or open an issue for discussion.

## 📄 License
This project is currently licensed under the MIT License. Please check the repository for more details.