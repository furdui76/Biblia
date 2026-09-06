using System.Windows;

namespace Biblia
{
    public partial class MainWindow : Window
    {
       

        public MainWindow()
        {
            InitializeComponent();
        }

        private void OpenBibliaWindow(object sender, RoutedEventArgs e)
        {
            new BibliaWindow().Show();
        }

       

        private void OpenTemeWindow(object sender, RoutedEventArgs e)
        {
            var f = new TemeWindow();
            f.ShowDialog();
        }

        private void OpenExplicatiiWindow(object sender, RoutedEventArgs e)
        {
            var fereastra = new ExplicatiiWindow();
            fereastra.ShowDialog();
        }


        private void OpenIstoricWindow(object sender, RoutedEventArgs e)
        {
            var fereastra = new FereastraIstoric();
            fereastra.Owner = this;
            fereastra.ShowDialog();
        }

        private void OpenArhivaWindow(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Fereastra Arhivă nu este încă implementată.");
        }
    }
}

