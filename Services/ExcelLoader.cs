using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using EuroSearchApp.Models;
using OfficeOpenXml; // EPPlus

namespace EuroSearchApp.Services
{
    public static class ExcelLoader
    {
        public static List<PersonRecord> Load(string filePath)
        {
            var output = new List<PersonRecord>();

            if (!File.Exists(filePath)) return output;

            try
            {
                // Ρύθμιση για το EPPlus (δωρεάν χρήση)
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var ws = package.Workbook.Worksheets.FirstOrDefault();
                    if (ws == null || ws.Dimension == null) return output;

                    // --- ΒΗΜΑ 1: Εντοπισμός Στηλών (Ψάχνουμε "έξυπνα" τις επικεφαλίδες) ---
                    int colComments = -1;
                    int colName = -1;
                    int colAfm = -1;
                    int colPhone1 = -1;
                    int colPhone2 = -1;
                    int colGift = -1;

                    // Σαρώνουμε την 1η γραμμή για να βρούμε πού είναι τι
                    for (int c = 1; c <= ws.Dimension.End.Column; c++)
                    {
                        var header = ws.Cells[1, c].Text.Trim();

                        // 1. ΕΥΡΕΣΗ ΟΝΟΜΑΤΟΣ
                        if (IsHeader(header, "Επωνυμία", "TITLE", "Name", "Ονοματεπώνυμο"))
                            colName = c;

                        // 2. ΕΥΡΕΣΗ ΑΦΜ (Εδώ ήταν το πρόβλημα πριν)
                        else if (IsHeader(header, "ΑΦΜ", "AFM", "Α.Φ.Μ.", "Α.Φ.Μ", "VAT", "Επαφές - Α.Φ.Μ"))
                            colAfm = c;

                        // 3. ΕΥΡΕΣΗ ΤΗΛΕΦΩΝΟΥ 1
                        else if (IsHeader(header, "Τηλέφωνο 1", "Τηλέφωνο", "TEL", "Phone", "Mobile", "Κινητό"))
                            colPhone1 = c;

                        // 4. ΕΥΡΕΣΗ ΤΗΛΕΦΩΝΟΥ 2
                        else if (IsHeader(header, "Τηλέφωνο 2", "Phone 2", "TEL 2"))
                            colPhone2 = c;

                        // 5. ΕΥΡΕΣΗ ΔΩΡΟΥ
                        else if (IsHeader(header, "ΔΩΡΟ", "Gift", "SelectedGift"))
                            colGift = c;

                        // 6. ΕΥΡΕΣΗ ΣΧΟΛΙΩΝ
                        else if (IsHeader(header, "Συμμετέχων", "Σχόλια", "Comments", "Παρατηρήσεις"))
                            colComments = c;
                    }

                    // --- ΒΗΜΑ 2: Διάβασμα Δεδομένων ---
                    for (int row = 2; row <= ws.Dimension.End.Row; row++)
                    {
                        // Αν δεν βρήκαμε στήλη ονόματος, δοκιμάζουμε την 2η ή την 1η ως λύση ανάγκης
                        if (colName == -1) colName = 2;

                        // Παίρνουμε το όνομα
                        var name = ws.Cells[row, colName].Value?.ToString()?.Trim();
                        if (string.IsNullOrWhiteSpace(name)) continue; // Αν δεν έχει όνομα, αγνόησε τη γραμμή

                        var p = new PersonRecord();
                        p.Επωνυμία = name;

                        // ΑΦΜ: Χρησιμοποιούμε .Text για να κρατήσουμε τα μηδενικά (π.χ. 099...)
                        if (colAfm != -1)
                            p.ΑΦΜ = ws.Cells[row, colAfm].Text?.Trim();

                        // Τηλέφωνα
                        if (colPhone1 != -1)
                            p.Τηλέφωνο = ws.Cells[row, colPhone1].Text?.Trim();

                        if (colPhone2 != -1)
                            p.Τηλέφωνο2 = ws.Cells[row, colPhone2].Text?.Trim();

                        // Σχόλια
                        if (colComments != -1)
                            p.Comments = ws.Cells[row, colComments].Value?.ToString()?.Trim();

                        // Δώρο
                        if (colGift != -1)
                        {
                            string gift = ws.Cells[row, colGift].Value?.ToString()?.Trim();
                            p.SelectedGift = gift;
                            if (!string.IsNullOrEmpty(gift)) p.Selected = true;
                        }

                        output.Add(p);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error loading Excel: " + ex.Message);
            }

            return output;
        }

        // Βοηθητική μέθοδος για να ελέγχουμε πολλές πιθανές ονομασίες επικεφαλίδων
        private static bool IsHeader(string header, params string[] candidates)
        {
            if (string.IsNullOrWhiteSpace(header)) return false;

            foreach (var candidate in candidates)
            {
                // Έλεγχος για ακριβή ισοτιμία
                if (string.Equals(header, candidate, StringComparison.OrdinalIgnoreCase)) return true;

                // Έλεγχος αν περιέχεται (π.χ. το "Επαφές - Α.Φ.Μ" περιέχει το "Α.Φ.Μ")
                if (candidate.Length > 2 && header.IndexOf(candidate, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            }
            return false;
        }
    }
}