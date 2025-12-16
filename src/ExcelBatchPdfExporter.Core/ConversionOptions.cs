namespace ExcelBatchPdfExporter.Core;

public record ConversionOptions(
    SheetSelectionMode Mode,
    bool OverwriteExisting
);
