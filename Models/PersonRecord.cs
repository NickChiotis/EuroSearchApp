using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EuroSearchApp.Models
{
    public class PersonRecord
    {
        // Αυτό λύνει το error: 'Selected' property not found
        public bool Selected { get; set; }

        public string Επωνυμία { get; set; }

        // Αυτό λύνει το error: 'ΑΦΜ' property not found
        public string ΑΦΜ { get; set; }

        // Αυτό λύνει το error: 'Τηλέφωνο' property not found
        public string Τηλέφωνο { get; set; }

        public string Τηλέφωνο2 { get; set; }
    }
}

