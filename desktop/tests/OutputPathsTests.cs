using System;
using System.IO;
using ExcelPdfConverter;
internal static class OutputPathsTests
{
    private static void Check(bool condition, string name) { if (!condition) throw new Exception(name); Console.WriteLine("PASS: " + name); }
    private static int Main()
    {
        string directory = Path.Combine(Path.GetTempPath(), "ExcelPdfTests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            Check(OutputPaths.IsSupported("BOOK.XLSX") && OutputPaths.IsSupported("book.xls") && OutputPaths.IsSupported("book.xlsm") && OutputPaths.IsSupported("book.xlsb"), "Excel extensions");
            Check(!OutputPaths.IsSupported("book.csv") && !OutputPaths.IsSupported("book.xlsx.exe"), "Unsupported extensions");
            Check(OutputPaths.Available(directory, "report") == Path.Combine(directory, "report.pdf"), "Initial name");
            File.WriteAllText(Path.Combine(directory, "report.pdf"), "existing");
            Directory.CreateDirectory(Path.Combine(directory, "report (1).pdf"));
            Check(OutputPaths.Available(directory, "report") == Path.Combine(directory, "report (2).pdf"), "Skip existing file and directory");
            File.WriteAllText(OutputPaths.Available(directory, "report"), "second");
            Check(OutputPaths.Available(directory, "report") == Path.Combine(directory, "report (3).pdf"), "Repeated conversion name");
            Check(File.ReadAllText(Path.Combine(directory, "report.pdf")) == "existing", "Existing file unchanged");
            return 0;
        }
        finally { Directory.Delete(directory, true); }
    }
}
