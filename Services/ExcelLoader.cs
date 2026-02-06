using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using EuroSearchApp.Models;
using ExcelDataReader;
using System.Text;
using System.Linq;

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
                            UseHeaderRow = true // Χρησιμοποιούμε την 1η γραμμή ως τίτλους
                        }
                    });

                    var table = result.Tables[0];

                    foreach (DataRow row in table.Rows)
                    {
                        // Διαβάζουμε πρώτα το δώρο για να αποφασίσουμε αν είναι Selected
                        string gift = Get(row, "ΔΩΡΟ");

                        var record = new PersonRecord
                        {
                            // 1. Σχόλια/Συμμετέχων από τη στήλη "Συμμετέχων"
                            Comments = Get(row, "Συμμετέχων"),

                            // 2. Επωνυμία
                            Επωνυμία = Get(row, "Επωνυμία"),

                            // 3. ΑΦΜ
                            ΑΦΜ = Get(row, "ΑΦΜ"),

                            // 4. Τηλέφωνο 1
                            Τηλέφωνο = Get(row, "Τηλέφωνο 1"),

                            // 5. Τηλέφωνο 2
                            Τηλέφωνο2 = Get(row, "Τηλέφωνο 2"),

                            // 6. Δώρο
                            SelectedGift = gift,

                            // Αυτόματη επιλογή αν υπάρχει δώρο
                            Selected = !string.IsNullOrWhiteSpace(gift)
                        };

                        list.Add(record);
                    }
                }
            }

            return list;
        }

        private static string Get(DataRow row, string columnName)
        {
            // Έλεγχος αν υπάρχει η στήλη για να μην "κρασάρει" το πρόγραμμα
            if (!row.Table.Columns.Contains(columnName)) return "";
            return row[columnName]?.ToString()?.Trim() ?? "";
        }
    }
}