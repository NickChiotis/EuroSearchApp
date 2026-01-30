using System;
using System.Diagnostics;
using System.IO;
using System.Windows;

namespace EuroSearchApp
{
    public static class ExcelAssetHelper
    {
        /// <summary>
        /// Extract ένα .xlsx από WPF Resource (Assets/...) σε temp και επιστρέφει το path.
        /// </summary>
        public static string ExtractExcelFromAssetsToTemp(string resourceRelativePath, string fileNamePrefix = "EuroSearch_Template")
        {
            // π.χ. resourceRelativePath = "Assets/Templates/Template.xlsx"
            var uri = new Uri($"pack://application:,,,/{resourceRelativePath}", UriKind.Absolute);

            var sri = Application.GetResourceStream(uri);
            if (sri == null)
                throw new FileNotFoundException(
                    $"Δεν βρέθηκε Resource: {resourceRelativePath}. " +
                    $"Έλεγξε ότι το αρχείο έχει Build Action = Resource.");

            string destPath = Path.Combine(
                Path.GetTempPath(),
                $"{fileNamePrefix}_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}.xlsx"
            );

            using (sri.Stream)
            using (var fs = new FileStream(destPath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                sri.Stream.CopyTo(fs);
            }

            return destPath;
        }

        /// <summary>
        /// Ανοίγει το αρχείο με το default app (Excel).
        /// </summary>
        public static void OpenWithDefaultApp(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException("Δεν βρέθηκε το αρχείο:", path);

            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
    }
}
