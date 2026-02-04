using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;

namespace EuroSearchApp.Models
{
    public class PersonRecord : INotifyPropertyChanged
    {
        // --- 1. SELECTED ---
        private bool _selected;
        public bool Selected
        {
            get { return _selected; }
            set
            {
                if (_selected != value)
                {
                    _selected = value;
                    OnPropertyChanged(nameof(Selected));
                }
            }
        }

        // --- 2. ΣΧΟΛΙΑ ---
        private string _comments;
        public string Comments
        {
            get { return _comments; }
            set
            {
                if (_comments != value)
                {
                    _comments = value;
                    OnPropertyChanged(nameof(Comments));
                }
            }
        }

        // --- 3. ΑΠΛΑ ΠΕΔΙΑ (Δεν χρειάζονται ειδική λογική ακόμα) ---
        public string Επωνυμία { get; set; }
        public string ΑΦΜ { get; set; }

        // --- 4. ΤΗΛΕΦΩΝΑ (ΒΕΛΤΙΣΤΟΠΟΙΗΜΕΝΑ) ---
        // Εδώ γίνεται η μαγεία για να μην κολλάει η αναζήτηση

        // Κρυφά πεδία (Backing fields)
        private string _τηλέφωνο;
        private string _τηλέφωνο2;

        // Πεδία που κρατάνε ΜΟΝΟ τα νούμερα (για γρήγορη αναζήτηση)
        public string CleanPhone1 { get; private set; }
        public string CleanPhone2 { get; private set; }

        public string Τηλέφωνο
        {
            get { return _τηλέφωνο; }
            set
            {
                if (_τηλέφωνο != value)
                {
                    _τηλέφωνο = value;
                    // Υπολογίζουμε το καθαρό νούμερο ΜΙΑ φορά, εδώ!
                    CleanPhone1 = NormalizeDigits(value);
                    OnPropertyChanged(nameof(Τηλέφωνο));
                }
            }
        }

        public string Τηλέφωνο2
        {
            get { return _τηλέφωνο2; }
            set
            {
                if (_τηλέφωνο2 != value)
                {
                    _τηλέφωνο2 = value;
                    // Υπολογίζουμε το καθαρό νούμερο ΜΙΑ φορά, εδώ!
                    CleanPhone2 = NormalizeDigits(value);
                    OnPropertyChanged(nameof(Τηλέφωνο2));
                }
            }
        }

        // --- ΒΟΗΘΗΤΙΚΗ ΜΕΘΟΔΟΣ ---
        // Πολύ γρήγορος καθαρισμός συμβόλων (κρατάει μόνο ψηφία)
        private static string NormalizeDigits(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";

            char[] buffer = new char[s.Length];
            int idx = 0;

            foreach (char c in s)
            {
                if (char.IsDigit(c))
                {
                    buffer[idx++] = c;
                }
            }
            return new string(buffer, 0, idx);
        }

        // --- INotifyPropertyChanged Implementation ---
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}