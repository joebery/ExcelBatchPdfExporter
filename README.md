# ExcelBatchPdfExporter

A Windows desktop application for batch converting Excel files to PDF using Excel’s native Print-to-PDF functionality.

The application is built with .NET 8 (WPF) and uses late-bound Excel COM automation to avoid Office Interop version issues while maintaining perfect layout fidelity.

---

## Features

- Batch convert multiple Excel files in one operation
- Global configuration to export a specific worksheet index across all files
- Automatic validation and warning for files missing the selected sheet
- Save PDFs next to input files or in a single chosen output folder
- Deterministic output naming:
