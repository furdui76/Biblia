using System.Text.RegularExpressions;
using System.Windows;

namespace Biblia
{
    public partial class EmailWindow : Window
    {
        public string Destinatar
        {
            get => txtDestinatar.Text;
            set => txtDestinatar.Text = value;
        }

        public string Subiect
        {
            get => txtSubiect.Text;
            set => txtSubiect.Text = value;
        }

        public string Continut
        {
            get => txtMesaj.Text;
            set => txtMesaj.Text = value;
        }


        public EmailWindow(string predicaInitiala)
        {
            InitializeComponent();
            txtMesaj.Text = predicaInitiala; // ✅ predica vine direct din fereastra principală
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            Destinatar = txtDestinatar.Text?.Trim() ?? string.Empty;
            Subiect = txtSubiect.Text?.Trim() ?? "Predica";
            Continut = txtMesaj.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(Destinatar))
            {
                MessageBox.Show("Te rog scrie adresa destinatarului.");
                return;
            }

            if (!Regex.IsMatch(Destinatar, @"^[^@\s]+@[^@\s]+\.[^@\s]+$"))
            {
                MessageBox.Show("Adresa de email nu este validă.");
                return;
            }

            DialogResult = true;
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}



