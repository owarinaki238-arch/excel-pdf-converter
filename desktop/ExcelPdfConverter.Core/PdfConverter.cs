using System;
using System.IO;
using System.Runtime.InteropServices;
namespace ExcelPdfConverter.Core
{
    public static class PdfConverter
    {
        // Call only on the worker STA thread. Never attach to an existing Excel session.
        public static string Convert(string source, string directory, bool ignorePrintArea)
        {
            object application = null, workbooks = null, workbook = null;
            string temporary = Path.Combine(directory, "." + Guid.NewGuid().ToString("N") + ".pdf");
            try
            {
                Type type = Type.GetTypeFromProgID("Excel.Application");
                if (type == null) throw new InvalidOperationException("デスクトップ版Microsoft Excelをインストールしてください。");
                application = Activator.CreateInstance(type);
                dynamic excel = application;
                excel.Visible = false;
                excel.DisplayAlerts = false;
                excel.EnableEvents = false;
                excel.AutomationSecurity = 3; // msoAutomationSecurityForceDisable
                excel.AskToUpdateLinks = false;
                workbooks = excel.Workbooks;
                dynamic books = workbooks;
                // Empty passwords prevent interactive password prompts for protected workbooks.
                workbook = books.Open(Filename: Path.GetFullPath(source), UpdateLinks: 0,
                    ReadOnly: true, Password: "", WriteResPassword: "",
                    IgnoreReadOnlyRecommended: true, Notify: false, AddToMru: false);
                dynamic book = workbook;
                book.ExportAsFixedFormat(Type: 0, Filename: temporary, Quality: 0,
                    IncludeDocProperties: true, IgnorePrintAreas: ignorePrintArea, OpenAfterPublish: false);
                using (var stream = File.OpenRead(temporary))
                {
                    byte[] header = new byte[5];
                    if (stream.Read(header, 0, 5) != 5 || System.Text.Encoding.ASCII.GetString(header) != "%PDF-")
                        throw new IOException("Excelが有効なPDFを出力しませんでした。");
                }
                // File.Move never overwrites another file, including one created concurrently.
                string destination = OutputPaths.Available(directory, Path.GetFileNameWithoutExtension(source));
                File.Move(temporary, destination);
                return destination;
            }
            finally
            {
                if (workbook != null)
                {
                    try { ((dynamic)workbook).Close(SaveChanges: false); } catch (COMException) { }
                    finally { Release(workbook); }
                }
                Release(workbooks);
                if (application != null)
                {
                    try { ((dynamic)application).Quit(); } catch (COMException) { }
                    finally { Release(application); }
                }
                try { if (File.Exists(temporary)) File.Delete(temporary); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }
        }
        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
        }
    }
}

