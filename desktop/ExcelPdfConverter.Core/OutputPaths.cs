using System;
using System.IO;
namespace ExcelPdfConverter.Core
{
    public static class OutputPaths
    {
        public static bool IsSupported(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            return ext == ".xlsx" || ext == ".xls" || ext == ".xlsm" || ext == ".xlsb";
        }
        public static string Available(string directory, string name)
        {
            for (int i = 0; ; i++)
            {
                string candidate = Path.Combine(directory, name + (i == 0 ? "" : " (" + i + ")") + ".pdf");
                if (!File.Exists(candidate) && !Directory.Exists(candidate)) return candidate;
            }
        }
    }
}

