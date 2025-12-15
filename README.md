# ExcelBatchPdfExporter

Batch convert Excel workbooks to PDF using Excel’s built-in
`ExportAsFixedFormat` (sheet-selectable) with a simple WPF GUI.

Windows-only — requires Microsoft Excel installed.

## Features
- Select multiple Excel files or an entire folder
- Export:
  - Active sheet
  - All sheets (one PDF per sheet)
  - Sheet by name
  - Sheet by index (1-based)
- Choose output folder
- Overwrite / skip existing PDFs
- Progress bar + log

## Tech
- .NET (C#)
- WPF
- Microsoft.Office.Interop.Excel

## Prerequisites
- Windows 10/11
- Microsoft Excel installed
- .NET 8 SDK (or 7+)

