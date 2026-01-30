using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using EuroSearchApp.Models;

namespace EuroSearchApp.Services
{
    public static class ExcelLoader
    {
        public static List<PersonRecord> Load(string filePath)
        {
            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet(1);

                var headerRow = ws.Row(1);
                int colName = FindColumn(headerRow, "Ονομα");
                int colPhone = FindColumn(headerRow, "Τηλέφωνο");
                int colAfm = FindColumn(headerRow, "ΑΦΜ");

                if (colName == -1 || colPhone == -1 || colAfm == -1)
                    throw new Exception("Δεν βρέθηκαν headers: Ονομα, Τηλέφωνο, ΑΦΜ (στην 1η γραμμή).");

                int lastRow = ws.LastRowUsed() != null ? ws.LastRowUsed().RowNumber() : 1;
                var list = new List<PersonRecord>();

                for (int r = 2; r <= lastRow; r++)
                {
                    var row = ws.Row(r);

                    string name = row.Cell(colName).GetString().Trim();
                    string phone = row.Cell(colPhone).GetString().Trim();
                    string afm = row.Cell(colAfm).GetString().Trim();

                    if (string.IsNullOrWhiteSpace(name) &&
                        string.IsNullOrWhiteSpace(phone) &&
                        string.IsNullOrWhiteSpace(afm))
                        continue;

                    list.Add(new PersonRecord
                    {
                        Ονομα = name,
                        Τηλέφωνο = phone,
                        ΑΦΜ = afm
                    });
                }

                return list;
            }
        }

        private static int FindColumn(IXLRow headerRow, string headerName)
        {
            foreach (var cell in headerRow.CellsUsed())
            {
                var text = cell.GetString().Trim();
                if (string.Equals(text, headerName, StringComparison.OrdinalIgnoreCase))
                    return cell.Address.ColumnNumber;
            }
            return -1;
        }
    }
}
