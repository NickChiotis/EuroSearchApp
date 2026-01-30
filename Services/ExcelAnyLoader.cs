using System;
using System.Data;
using System.Linq;
using ClosedXML.Excel;

namespace EuroSearchApp.Services
{
    public static class ExcelAnyLoader
    {
        public static DataTable LoadAllToDataTable(string path, bool firstRowHasHeaders)
        {
            var dt = new DataTable();
            dt.Columns.Add("Επιλογή", typeof(bool));

            using (var wb = new XLWorkbook(path))
            {
                var ws = wb.Worksheets.FirstOrDefault();
                if (ws == null) return dt;

                var range = ws.RangeUsed();
                if (range == null) return dt;

                int firstRow = range.FirstRowUsed().RowNumber();
                int lastRow = range.LastRowUsed().RowNumber();
                int firstCol = range.FirstColumnUsed().ColumnNumber();
                int lastCol = range.LastColumnUsed().ColumnNumber();

                // headers ή Column1..
                for (int c = firstCol; c <= lastCol; c++)
                {
                    string colName = "";

                    if (firstRowHasHeaders)
                        colName = (ws.Cell(firstRow, c).GetString() ?? "").Trim();

                    if (string.IsNullOrWhiteSpace(colName))
                        colName = "Column" + (c - firstCol + 1);

                    colName = MakeUnique(dt, colName);
                    dt.Columns.Add(colName, typeof(string));
                }

                int startDataRow = firstRowHasHeaders ? firstRow + 1 : firstRow;

                for (int r = startDataRow; r <= lastRow; r++)
                {
                    var dr = dt.NewRow();
                    dr["Επιλογή"] = false;

                    int idx = 1; // 0 = Επιλογή
                    for (int c = firstCol; c <= lastCol; c++)
                    {
                        dr[idx] = ws.Cell(r, c).GetFormattedString();
                        idx++;
                    }

                    dt.Rows.Add(dr);
                }
            }

            return dt;
        }

        private static string MakeUnique(DataTable dt, string baseName)
        {
            string name = baseName;
            int i = 1;
            while (dt.Columns.Contains(name))
            {
                i++;
                name = baseName + "_" + i;
            }
            return name;
        }
    }
}