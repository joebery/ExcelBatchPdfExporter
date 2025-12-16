using System.Collections.ObjectModel;
using System.Windows;
using ExcelBatchPdfExporter.Core;

namespace ExcelBatchPdfExporter.Gui;

public partial class MainWindow : Window
{
    private readonly ExcelToPdfConverter _converter = new();

    private readonly ObservableCollection<string> _files = new();

    private string? _outputFolder; // chosen via button

    public MainWindow()
    {
        InitializeComponent();

        FilesList.ItemsSource = _files;

        StatusText.Text = "Add files, set sheet index, then Convert ALL.";
        Log("Ready.");
    }

    private void AddFiles_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Title = "Select Excel/CSV file(s)",
            Filter = "Supported Files (*.xls;*.xlsx;*.xlsm;*.xlsb;*.csv)|*.xls;*.xlsx;*.xlsm;*.xlsb;*.csv|All Files|*.*",
            Multiselect = true
        };

        if (dlg.ShowDialog() == true)
        {
            foreach (var f in dlg.FileNames)
                AddFileIfMissing(f);

            Log($"Added {dlg.FileNames.Length} file(s).");
            StatusText.Text = $"Files in list: {_files.Count}";
        }
    }

    private void RemoveFiles_Click(object sender, RoutedEventArgs e)
    {
        var selected = FilesList.SelectedItems.Cast<string>().ToList();
        foreach (var f in selected)
            _files.Remove(f);

        StatusText.Text = $"Files in list: {_files.Count}";
        Log("Removed selected file(s).");
    }

    private void ClearFiles_Click(object sender, RoutedEventArgs e)
    {
        _files.Clear();
        StatusText.Text = "Cleared file list.";
        Log("Cleared file list.");
    }

    private void BrowseOutput_Click(object sender, RoutedEventArgs e)
    {
        using var dlg = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = "Select output folder",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true
        };

        if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            _outputFolder = dlg.SelectedPath;
            OutputFolderText.Text = _outputFolder;
            Log($"Output folder set: {_outputFolder}");
        }
    }

    private void InspectFirstFile_Click(object sender, RoutedEventArgs e)
    {
        if (_files.Count == 0)
        {
            Log("No files in list to inspect.");
            return;
        }

        var first = _files[0];
        try
        {
            Log($"Inspecting first file: {first}");
            var sheets = _converter.GetWorksheetNames(first);
            Log($"Sheet count: {sheets.Count}");
            for (int i = 0; i < sheets.Count; i++)
                Log($"  [{i + 1}] {sheets[i]}");
        }
        catch (Exception ex)
        {
            Log(ex.ToString());
        }
    }

    private void ConvertAll_Click(object sender, RoutedEventArgs e)
    {
        if (_files.Count == 0)
        {
            Log("No files in list to convert.");
            return;
        }

        if (!int.TryParse(SheetIndexBox.Text.Trim(), out var sheetIndex) || sheetIndex < 1)
        {
            Log("Global sheet index must be a number >= 1.");
            return;
        }

        var overwrite = OverwriteCheck.IsChecked == true;

        // Decide output mode:
        // - if UseOutputFolderCheck is checked: must have chosen folder
        // - otherwise: save next to each input file
        var useOutputFolder = UseOutputFolderCheck.IsChecked == true;
        if (useOutputFolder && string.IsNullOrWhiteSpace(_outputFolder))
        {
            Log("You enabled 'Use output folder' but haven't chosen one yet.");
            return;
        }

        StatusText.Text = "Converting...";
        Log($"Batch converting {_files.Count} file(s) using sheet index {sheetIndex}...");

        var missing = new List<string>();
        var okCount = 0;

        foreach (var file in _files)
        {
            try
            {
                var outputDir = useOutputFolder
                    ? _outputFolder!
                    : (System.IO.Path.GetDirectoryName(file) ?? Environment.CurrentDirectory);

                var folderName = System.IO.Path.GetFileName(System.IO.Path.GetDirectoryName(file) ?? "") ?? "output";
                var excelName = System.IO.Path.GetFileNameWithoutExtension(file);

                // Required naming: <foldername>-<excelname>.pdf
                var pdfName = $"{MakeSafe(folderName)}-{MakeSafe(excelName)}.pdf";
                var pdfPath = System.IO.Path.Combine(outputDir, pdfName);

                Log($"Converting: {file}");
                _converter.ExportSheetByIndexToPdf(file, sheetIndex, pdfPath, overwrite);
                Log($"Saved: {pdfPath}");
                okCount++;
            }
            catch (Exception ex)
            {
                // If sheet index doesn't exist, log and track missing
                var msg = ex.Message ?? "";
                if (msg.Contains("Sheet index", StringComparison.OrdinalIgnoreCase) &&
                    msg.Contains("not found", StringComparison.OrdinalIgnoreCase))
                {
                    missing.Add(System.IO.Path.GetFileName(file));
                    Log($"MISSING SHEET: {file} -> {msg}");
                }
                else
                {
                    Log($"ERROR converting {file}: {ex}");
                }
            }
        }

        StatusText.Text = $"Done. Converted {okCount}/{_files.Count}.";

        if (missing.Count > 0)
        {
            var text = $"These files do not contain sheet index {sheetIndex}:\n\n" +
                       string.Join("\n", missing);

            // Explicit WPF MessageBox to avoid ambiguity with WinForms MessageBox
            System.Windows.MessageBox.Show(text, "Missing sheet", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void AddFileIfMissing(string path)
    {
        if (!_files.Contains(path, StringComparer.OrdinalIgnoreCase))
            _files.Add(path);
    }

    private static string MakeSafe(string name)
    {
        foreach (var c in System.IO.Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name.Trim();
    }

    private void Log(string msg)
    {
        LogBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}{Environment.NewLine}");
        LogBox.ScrollToEnd();
    }
}
