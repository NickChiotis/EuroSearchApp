using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace EuroSearchApp.Models
{
    public class PersonRecord : INotifyPropertyChanged
    {
        // --- 1. SELECTED (CHECKBOX) ---
        private bool _selected;
        public bool Selected
        {
            get { return _selected; }
            set
            {
                if (_selected != value)
                {
                    _selected = value;
                    OnPropertyChanged();
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
                    OnPropertyChanged();
                }
            }
        }

        // --- 3. ΒΑΣΙΚΑ ΠΕΔΙΑ (ΕΠΩΝΥΜΙΑ & ΑΦΜ) ---
        // Προσοχή: Εδώ βάλαμε OnPropertyChanged για να ενημερώνεται το UI
        // αυτόματα μόλις έρθουν τα στοιχεία από την ΑΑΔΕ.

        private string _eponymia;
        public string Επωνυμία
        {
            get { return _eponymia; }
            set
            {
                if (_eponymia != value)
                {
                    _eponymia = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _afm;
        public string ΑΦΜ
        {
            get { return _afm; }
            set
            {
                if (_afm != value)
                {
                    _afm = value;
                    OnPropertyChanged();
                }
            }
        }

        // Προαιρετικό: Αν έχεις ID στη βάση δεδομένων για το SQL Update
        public int IDinDB { get; set; }

        // --- 4. ΤΗΛΕΦΩΝΑ (PERFORMANCE OPTIMIZATION) ---
        // Εδώ γίνεται ο αυτόματος καθαρισμός για να μην κολλάει η αναζήτηση.

        private string _tilefono;
        private string _tilefono2;

        // Αυτά τα πεδία χρησιμοποιεί το φίλτρο (Μόνο νούμερα)
        public string CleanPhone1 { get; private set; }
        public string CleanPhone2 { get; private set; }

        private string _selectedGift;
        public string SelectedGift
        {
            get
            {
                // Αν είναι null ή κενό, επέστρεψε το placeholder
                return string.IsNullOrEmpty(_selectedGift) ? "Επιλογή Δώρου" : _selectedGift;
            }
            set
            {
                _selectedGift = value;
                OnPropertyChanged(nameof(SelectedGift));
            }
        } // Εδώ θα αποθηκεύεται η επιλογή από το dropdown

        public string Τηλέφωνο
        {
            get { return _tilefono; }
            set
            {
                if (_tilefono != value)
                {
                    _tilefono = value;
                    // Υπολογισμός καθαρού αριθμού ΜΙΑ φορά κατά την ανάθεση
                    CleanPhone1 = NormalizeDigits(value);
                    OnPropertyChanged();
                }
            }
        }

        public string Τηλέφωνο2
        {
            get { return _tilefono2; }
            set
            {
                if (_tilefono2 != value)
                {
                    _tilefono2 = value;
                    // Υπολογισμός καθαρού αριθμού ΜΙΑ φορά κατά την ανάθεση
                    CleanPhone2 = NormalizeDigits(value);
                    OnPropertyChanged();
                }
            }
        }

        // --- 5. ΔΙΕΥΘΥΝΣΗ (Προαιρετικό, για την ΑΑΔΕ) ---
        private string _address;
        public string Διεύθυνση
        {
            get { return _address; }
            set
            {
                if (_address != value)
                {
                    _address = value;
                    OnPropertyChanged();
                }
            }
        }


        // --- HELPER: ΓΡΗΓΟΡΟΣ ΚΑΘΑΡΙΣΜΟΣ STRING ---
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

        // Η μέθοδος αυτή ειδοποιεί το UI ότι άλλαξε κάποια τιμή
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}