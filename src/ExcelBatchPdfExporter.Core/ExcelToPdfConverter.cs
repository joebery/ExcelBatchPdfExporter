using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelBatchPdfExporter.Core;

public sealed class ExcelToPdfConverter
{
    /// <summary>
    /// Reads worksheet names from a workbook (late-bound COM).
    /// Useful for inspecting the first file to confirm sheet counts/names.
    /// </summary>
    public IReadOnlyList<string> GetWorksheetNames(string excelPath)
    {
        if (!File.Exists(excelPath))
            throw new FileNotFoundException("Excel file not found.", excelPath);

        object? app = null;
        object? workbooks = null;
        object? workbook = null;
        object? sheets = null;
        Type? excelType = null;

        var names = new List<string>();

        try
        {
            excelType = Type.GetTypeFromProgID("Excel.Application")
                        ?? throw new InvalidOperationException("Excel.Application ProgID not found. Is Excel installed?");

            app = Activator.CreateInstance(excelType)
                  ?? throw new InvalidOperationException("Failed to create Excel.Application instance.");

            excelType.InvokeMember("Visible", BindingFlags.SetProperty, null, app, new object[] { false });
            excelType.InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, app, new object[] { false });

            workbooks = excelType.InvokeMember("Workbooks", BindingFlags.GetProperty, null, app, null)
                        ?? throw new InvalidOperationException("Failed to get Workbooks collection.");

            workbook = OpenReadOnly(workbooks, excelPath);

            sheets = workbook.GetType().InvokeMember("Sheets", BindingFlags.GetProperty, null, workbook, null)
                     ?? throw new InvalidOperationException("Failed to get Sheets collection.");

            var countObj = sheets.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, sheets, null);
            var count = Convert.ToInt32(countObj);

            for (int i = 1; i <= count; i++)
            {
                object? sheet = null;
                try
                {
                    sheet = sheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, sheets, new object[] { i });
                    if (sheet == null) continue;

                    var sheetName = (string)(sheet.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, sheet, null) ?? "");
                    if (!string.IsNullOrWhiteSpace(sheetName))
                        names.Add(sheetName);
                }
                finally
                {
                    if (sheet != null) ReleaseCom(sheet);
                }
            }

            return names;
        }
        finally
        {
            if (sheets != null) ReleaseCom(sheets);

            CloseAndReleaseWorkbook(workbook);
            if (workbooks != null) ReleaseCom(workbooks);

            QuitAndReleaseApp(app, excelType);

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    /// <summary>
    /// Exports a specific worksheet by 1-based index to PDF.
    /// Output file name must be passed in (caller controls naming scheme).
    /// Throws if sheet index doesn't exist.
    /// </summary>
    public void ExportSheetByIndexToPdf(string excelPath, int sheetIndex1Based, string pdfPath, bool overwrite)
    {
        if (!File.Exists(excelPath))
            throw new FileNotFoundException("Excel file not found.", excelPath);

        if (sheetIndex1Based < 1)
            throw new ArgumentOutOfRangeException(nameof(sheetIndex1Based), "Sheet index must be >= 1.");

        var outDir = Path.GetDirectoryName(pdfPath);
        if (!string.IsNullOrWhiteSpace(outDir))
            Directory.CreateDirectory(outDir);

        object? app = null;
        object? workbooks = null;
        object? workbook = null;
        object? sheets = null;
        object? sheet = null;

        Type? excelType = null;

        try
        {
            excelType = Type.GetTypeFromProgID("Excel.Application")
                        ?? throw new InvalidOperationException("Excel.Application ProgID not found. Is Excel installed?");

            app = Activator.CreateInstance(excelType)
                  ?? throw new InvalidOperationException("Failed to create Excel.Application instance.");

            excelType.InvokeMember("Visible", BindingFlags.SetProperty, null, app, new object[] { false });
            excelType.InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, app, new object[] { false });

            workbooks = excelType.InvokeMember("Workbooks", BindingFlags.GetProperty, null, app, null)
                        ?? throw new InvalidOperationException("Failed to get Workbooks collection.");

            workbook = OpenReadOnly(workbooks, excelPath);

            sheets = workbook.GetType().InvokeMember("Sheets", BindingFlags.GetProperty, null, workbook, null)
                     ?? throw new InvalidOperationException("Failed to get Sheets collection.");

            // Get sheet by 1-based index
            try
            {
                sheet = sheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, sheets, new object[] { sheetIndex1Based });
            }
            catch
            {
                sheet = null;
            }

            if (sheet == null)
                throw new InvalidOperationException($"Sheet index {sheetIndex1Based} not found.");

            // Activate just in case
            TryInvoke(sheet, "Activate");

            ExportSheetToPdf(sheet, pdfPath, overwrite);
        }
        finally
        {
            if (sheet != null) ReleaseCom(sheet);
            if (sheets != null) ReleaseCom(sheets);

            CloseAndReleaseWorkbook(workbook);
            if (workbooks != null) ReleaseCom(workbooks);

            QuitAndReleaseApp(app, excelType);

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    // ---------------- helpers ----------------

    private static object OpenReadOnly(object workbooks, string excelPath)
    {
        return workbooks.GetType().InvokeMember(
            "Open",
            BindingFlags.InvokeMethod,
            null,
            workbooks,
            new object[]
            {
                excelPath,
                Type.Missing, // UpdateLinks
                true,         // ReadOnly
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
            }) ?? throw new InvalidOperationException("Failed to open workbook.");
    }

    private static void ExportSheetToPdf(object sheet, string pdfPath, bool overwrite)
    {
        if (File.Exists(pdfPath))
        {
            if (!overwrite) return;
            File.Delete(pdfPath);
        }

        const int xlTypePDF = 0;
        const int xlQualityStandard = 0;

        sheet.GetType().InvokeMember(
            "ExportAsFixedFormat",
            BindingFlags.InvokeMethod,
            null,
            sheet,
            new object[]
            {
                xlTypePDF,
                pdfPath,
                xlQualityStandard,
                true,
                false,
                Type.Missing,
                Type.Missing,
                false,
                Type.Missing
            });

        if (!File.Exists(pdfPath))
            throw new IOException("ExportAsFixedFormat returned, but the PDF was not created.");
    }

    private static void TryInvoke(object comObj, string methodName)
    {
        try
        {
            comObj.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, null, comObj, Array.Empty<object>());
        }
        catch { /* ignore */ }
    }

    private static void CloseAndReleaseWorkbook(object? workbook)
    {
        if (workbook == null) return;

        try
        {
            workbook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, workbook, new object[] { false });
        }
        catch { /* ignore */ }

        ReleaseCom(workbook);
    }

    private static void QuitAndReleaseApp(object? app, Type? excelType)
    {
        if (app == null || excelType == null) return;

        try
        {
            excelType.InvokeMember("Quit", BindingFlags.InvokeMethod, null, app, null);
        }
        catch { /* ignore */ }

        ReleaseCom(app);
    }

    private static void ReleaseCom(object comObj)
    {
        try
        {
            if (Marshal.IsComObject(comObj))
                Marshal.FinalReleaseComObject(comObj);
        }
        catch { /* ignore */ }
    }
}
