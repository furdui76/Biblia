using System.Windows;

namespace Biblia
{
    public partial class FereastraEditareComentariu : Window
    {
        public string ComentariuFinal { get; private set; }

        public FereastraEditareComentariu(string comentariuInitial)
        {
            InitializeComponent();
            txtComentariu.Text = comentariuInitial;
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            ComentariuFinal = txtComentariu.Text;
            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}

