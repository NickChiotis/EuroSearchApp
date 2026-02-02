using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using ClosedXML.Excel;
using EuroSearchApp.Models;
using ExcelDataReader;
using System.Text;

namespace EuroSearchApp.Services
{
    public static class ExcelLoader
    {
        public static List<PersonRecord> Load(string path)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var list = new List<PersonRecord>();

            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = true
                        }
                    });

                    // Παίρνουμε το 1ο φύλλο
                    var table = result.Tables[0];

                    foreach (DataRow row in table.Rows)
                    {
                        var record = new PersonRecord
                        {
                            Selected = false,

                            // ΕΔΩ ΓΙΝΕΤΑΙ Η "ΜΑΓΕΙΑ" ΤΗΣ ΑΝΤΙΣΤΟΙΧΙΣΗΣ
                            // Αριστερά: Η κλάση μας | Δεξιά: Η στήλη στο Excel
                            Επωνυμία = Get(row, "Επωνυμία"),
                            ΑΦΜ = Get(row, "Επαφές - Α.Φ.Μ"),
                            Τηλέφωνο = Get(row, "Τηλέφωνο 1"),
                            Τηλέφωνο2 = Get(row, "Τηλέφωνο 2")
                        };

                        list.Add(record);
                    }
                }
            }

            return list;
        }

        private static string Get(DataRow row, string columnName)
        {
            if (!row.Table.Columns.Contains(columnName)) return "";
            return row[columnName]?.ToString()?.Trim() ?? "";
        }
    }
}
