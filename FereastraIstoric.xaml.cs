using Microsoft.Web.WebView2.Core;
using System.Windows;

namespace Biblia
{
    public partial class FereastraIstoric : Window
    {
        public FereastraIstoric()
        {
            InitializeComponent();
            Loaded += FereastraIstoric_Loaded;
        }

        private async void FereastraIstoric_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                await webViewTimeline.EnsureCoreWebView2Async();
                webViewTimeline.CoreWebView2.Navigate("https://ebiblia.ro/app/#_");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("❌ Nu pot încărca cronologia. Verifică conexiunea la internet și WebView2 Runtime.\n\nDetalii: " + ex.Message);
            }
        }
    }
}


