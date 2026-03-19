using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace EuroSearchApp
{
    /// <summary>
    /// Interaction logic for LauncherWindow.xaml
    /// </summary>
    public partial class LauncherWindow : Window
    {
        public LauncherWindow()
        {
            InitializeComponent();
        }

        private void OpenEuroWheel_Click(object sender, RoutedEventArgs e)
        {
            MainWindow wheelWindow = new MainWindow();
            wheelWindow.Show();
            this.Close(); // Κλείνει το μενού επιλογής
        }

        private void BtnEuroStatement_Click(object sender, RoutedEventArgs e)
        {
            StatementDashboard dashboard = new StatementDashboard();
            dashboard.Owner = this; // Για να μένει μπροστά από το κεντρικό παράθυρο
            dashboard.ShowDialog();
        }
    }
}
