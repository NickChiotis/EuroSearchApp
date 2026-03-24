using System.Windows;

namespace EuroSearchApp
{
    public partial class SignatureWindow : Window
    {
        public SignatureWindow()
        {
            InitializeComponent();
        }

        private void BtnFinish_Click(object sender, RoutedEventArgs e)
        {
            if (SignCanvas.Strokes.Count == 0)
            {
                MessageBox.Show("Παρακαλώ βάλτε την υπογραφή σας.");
                return;
            }
            this.DialogResult = true; // Προχωράμε
        }
    }
}