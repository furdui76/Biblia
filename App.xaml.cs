using PdfSharp.Fonts;
using System.Configuration;
using System.Data;
using System.Windows;

namespace Biblia
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            GlobalFontSettings.FontResolver = new CustomFontResolver();
            base.OnStartup(e);
        }



    }

}
