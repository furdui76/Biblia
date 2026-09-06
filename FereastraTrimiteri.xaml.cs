using Biblia;
using Biblia.Models;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Threading;


namespace Biblia
{
    
    
    
    public partial class FereastraTrimiteri : Window
    {
        private BibliaWindow bibliaWindow;
        private List<int> verseteDeEvidentiat = new List<int>();
        private BibliaWindow parent;

        public FereastraTrimiteri(BibliaWindow parent, string carte, int capitol, List<int> versete = null)
        {
            InitializeComponent();
            bibliaWindow = parent; // ✅ acum există
            this.parent = parent;
            titlu.Text = $"{carte} {capitol}";
            verseteDeEvidentiat = versete ?? new List<int>();
            IncarcaVersete(carte, capitol);
        }


        // Model pentru schiță
        

        // Stare globală a aplicației
        public static class AppState
        {
            public static bool ModSchitaActiv { get; set; } = false;
            public static List<VersetSchita> SchitaCurenta { get; set; } = new List<VersetSchita>();
            

        }

        private void IncarcaVersete(string carte, int capitol)
        {
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "biblia.db");
            using var conn = new SQLiteConnection($"Data Source={path}");
            conn.Open();

            var cmd = new SQLiteCommand(
                "SELECT Verset, Text FROM Versete WHERE Carte = @carte AND Capitol = @capitol ORDER BY Verset", conn);
            cmd.Parameters.AddWithValue("@carte", carte);
            cmd.Parameters.AddWithValue("@capitol", capitol);

            TextBlock versetTarget = null;
            listaVersete.Children.Clear();

            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                int nr = reader.GetInt32(0);
                string txt = reader.GetString(1);

                var wrap = new WrapPanel
                {
                    Orientation = Orientation.Horizontal,
                    Margin = new Thickness(0, 1, 0, 1),
                    VerticalAlignment = VerticalAlignment.Top
                };

                wrap.Loaded += (s, e) =>
                {
                    wrap.MaxWidth = Math.Max(100, listaVersete.ActualWidth - 20);
                };

                var tb = new TextBlock
                {
                    FontSize = 16,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 0, 0, 10),
                    FontWeight = FontWeights.Bold   // 🔥 AICI devine bold
                };


                tb.Inlines.Add(new Run($"{nr}. ") { FontWeight = FontWeights.Bold });

                int index = 0;
                while (index < txt.Length)
                {
                    int startRosu = txt.IndexOf("<ROSU>", index);
                    if (startRosu == -1)
                    {
                        tb.Inlines.Add(new Run(txt.Substring(index)));
                        break;
                    }

                    if (startRosu > index)
                        tb.Inlines.Add(new Run(txt.Substring(index, startRosu - index)));

                    int endRosu = txt.IndexOf("</ROSU>", startRosu);
                    if (endRosu == -1)
                    {
                        string rosuText = txt.Substring(startRosu + 6);
                        tb.Inlines.Add(new Run(rosuText) { Foreground = Brushes.Red });
                        break;
                    }

                    string rosuTextFinal = txt.Substring(startRosu + 6, endRosu - (startRosu + 6));
                    tb.Inlines.Add(new Run(rosuTextFinal) { Foreground = Brushes.Red });

                    index = endRosu + 7;
                }

                if (verseteDeEvidentiat.Contains(nr))
                {
                    tb.Background = Brushes.LightYellow;
                    tb.FontWeight = FontWeights.Bold;
                    versetTarget ??= tb;

                    // ✅ adaugă textul
                    wrap.Children.Add(tb);

                    if (bibliaWindow.EsteModSchitaActiv)
                    {
                        bool esteInSchita = AppState.SchitaCurenta.Any(v =>
                            v.Carte == carte &&
                            v.Capitol == capitol &&
                            v.Verset == nr);

                        var btnAdauga = new Button
                        {
                            Content = esteInSchita ? "✅" : "➕",
                            FontSize = 12,
                            Width = 20,
                            Height = 20,
                            Margin = new Thickness(6, 0, 0, 0),
                            VerticalAlignment = VerticalAlignment.Center,
                            IsEnabled = !esteInSchita,
                            Tag = new VersetSchita { Carte = carte, Capitol = capitol, Verset = nr, Text = txt }
                        };

                        // ToolTip inițial
                        btnAdauga.ToolTip = new ToolTip
                        {
                            Background = Brushes.LightYellow,
                            Padding = new Thickness(6),
                            Content = new TextBlock
                            {
                                Text = esteInSchita ? "Versetul este deja în schiță" : "Adaugă la schiță",
                                Foreground = esteInSchita ? Brushes.Green : Brushes.DarkBlue,
                                FontSize = 12,
                                FontWeight = FontWeights.Bold
                            }
                        };

                        // timpii pentru afișare instant
                        ToolTipService.SetInitialShowDelay(btnAdauga, 0);
                        ToolTipService.SetShowDuration(btnAdauga, 5000);
                        ToolTipService.SetBetweenShowDelay(btnAdauga, 0);

                        if (!esteInSchita)
                            btnAdauga.Click += BtnAdaugaLaSchita_Click;

                        // ❌ NU mai faci wrap.Children.Add(btnAdauga)
                        // ✅ inserezi butonul direct la finalul textului
                        tb.Inlines.Add(new InlineUIContainer(btnAdauga));
                    }
                }

                else
                {
                    wrap.Children.Add(tb); // ✅ verset normal
                }

                listaVersete.Children.Add(wrap);
            }

            if (versetTarget != null)
            {
                versetTarget.Loaded += (s, e) =>
                {
                    versetTarget.BringIntoView();
                    CenterElementInScrollViewer(versetTarget);
                };
            }
        }



        // 🔹 Handler pentru adăugare în schiță
        // 🔹 Handler pentru adăugare în schiță
        private void BtnAdaugaLaSchita_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is VersetSchita v)
            {
                parent.AdaugaLaSchita(v);

                btn.Content = "✅";
                btn.IsEnabled = true;

                // confirmare imediată cu Popup (toast)
                var popup = new Popup
                {
                    PlacementTarget = btn,
                    Placement = PlacementMode.Right,
                    StaysOpen = false,
                    Child = new Border   // atenție: Border, nu WpfBorder
                    {
                        Background = Brushes.LightYellow,
                        Padding = new Thickness(6),
                        Child = new TextBlock
                        {
                            Text = "Versetul a fost adăugat în schiță",
                            Foreground = Brushes.Green,
                            FontSize = 12,
                            FontWeight = FontWeights.Bold
                        }
                    }
                };

                popup.IsOpen = true;

                var timer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromSeconds(2)
                };
                timer.Tick += (s, args) =>
                {
                    popup.IsOpen = false;
                    timer.Stop();

                    // ToolTip pentru hover ulterior
                    btn.ToolTip = new ToolTip
                    {
                        Background = Brushes.LightYellow,   // fundalul ToolTip-ului
                        Padding = new Thickness(6),
                        Content = new TextBlock
                        {
                            Text = "Versetul este deja în schiță", // mesajul
                            Foreground = Brushes.Green,            // 👉 text verde
                            FontSize = 12,
                            FontWeight = FontWeights.Bold          // 👉 bold
                        }
                    };

                    // timpii pentru afișare instant
                    ToolTipService.SetInitialShowDelay(btn, 0);   // apare imediat
                    ToolTipService.SetShowDuration(btn, 5000);   // rămâne 5 secunde
                    ToolTipService.SetBetweenShowDelay(btn, 100);
                };
                timer.Start();
            }
        }






        private void CenterElementInScrollViewer(FrameworkElement element)
        {
            var scrollViewer = FindVisualChild<ScrollViewer>(this); // schimbat din listaVersete în this
            if (scrollViewer == null) return;

            var transform = element.TransformToAncestor(scrollViewer);
            var position = transform.Transform(new Point(0, 0));

            double offset = position.Y - (scrollViewer.ViewportHeight / 2) + (element.ActualHeight / 2);
            scrollViewer.ScrollToVerticalOffset(offset);
        }


        private T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T tChild)
                    return tChild;

                var result = FindVisualChild<T>(child);
                if (result != null)
                    return result;
            }
            return null;
        }

    }
}



