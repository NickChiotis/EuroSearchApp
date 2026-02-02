using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;

namespace EuroSearchApp.Models
{
    // Προσθέσαμε το ": INotifyPropertyChanged" δίπλα στο όνομα της κλάσης
    public class PersonRecord : INotifyPropertyChanged
    {
        // --- Η ΙΔΙΟΤΗΤΑ SELECTED (ΑΛΛΑΓΜΕΝΗ) ---
        // Την αλλάξαμε για να ειδοποιεί όταν τικάρεται/ξετικάρεται
        private bool _selected;
        public bool Selected
        {
            get { return _selected; }
            set
            {
                if (_selected != value)
                {
                    _selected = value;
                    // Αυτή η εντολή λέει στο σύστημα: "Εί, άλλαξε το Selected! Κάνε τα κουμάντα σου!"
                    OnPropertyChanged(nameof(Selected));
                }
            }
        }

        // --- ΤΑ ΥΠΟΛΟΙΠΑ ΠΕΔΙΑ ---
        // Αυτά μπορούν να μείνουν απλά (όπως τα είχες), γιατί δεν τα αλλάζεις εσύ στο Grid
        public string Επωνυμία { get; set; }

        public string ΑΦΜ { get; set; }

        public string Τηλέφωνο { get; set; }

        public string Τηλέφωνο2 { get; set; }


        // --- ΚΩΔΙΚΑΣ ΕΙΔΟΠΟΙΗΣΗΣ (BOILERPLATE) ---
        // Αυτό είναι στάνταρ κώδικας που χρειάζεται πάντα το WPF
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}

