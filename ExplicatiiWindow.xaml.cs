using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Media;
using System.Net;
namespace Biblia
{
    public partial class ExplicatiiWindow : Window
    {
        private static readonly string CONNECTION_STRING =
     Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "biblia.db");


        private readonly Dictionary<string, string> mapareCanonica = new()
        {
            { "Geneza", "GENEZA" }, { "Gen.", "GENEZA" }, { "Ex.", "EXODUL" },{ "Exod", "EXODUL" }, { "Levitic", "LEVITICUL" }, { "Numeri", "NUMERE" },
            { "Num.", "NUMERI" },
            { "Deuteronom", "DEUTERONOMUL" }, { "Iosua", "IOSUA" }, { "Judecători", "JUDECĂTORI" }, { "Rut", "RUT" },
            { "1 Samuel", "1 SAMUEL" }, { "2 Samuel", "2 SAMUEL" }, { "1 Împărați", "1 ÎMPĂRAȚI" }, { "2 Împărați", "2 ÎMPĂRAȚI" },
            { "1 Cronici", "1 CRONICI" }, { "2 Cronici", "2 CRONICI" }, { "Ezra", "EZRA" }, { "Neemia", "NEEMIA" },
            { "Estera", "ESTERA" }, { "Iov", "IOV" }, { "Psalmii", "PSALMI" }, { "Ps.", "PSALMII" },
            { "Proverbe", "PROVERBE" }, { "Eclesiastul", "ECLESIASTUL" }, { "Ecl.", "ECLESIASTUL" }, { "Cântarea Cântărilor", "CÂNTAREA CÂNTĂRILOR" },
            { "Is.", "ISAIA" }, { "Ier.", "IEREMIA" }, { "Plângerile", "PLÂNGERILE" }, { "Ezechiel", "EZECHEL" },{ "Ezec.", "EZECHEL" },
            { "Daniel", "DANIEL" }, { "Osea", "OSEA" }, { "Ioel", "IOEL" }, { "Amos", "AMOS" }, { "Obadia", "OBADIA" },
            { "Iona", "IONA" }, { "Mica", "MICA" }, { "Naum", "NAUM" }, { "Habacuc", "HABACUC" }, { "Țefania", "ȚEFANIA" },
            { "Hagai", "HAGAI" }, { "Zaharia", "ZAHARIA" }, { "Maleahi", "MALEAHI" }, { "Matei", "MATEI" }, { "Mt.", "MATEI" },
            { "Marcu", "MARCU" }, { "Mc.", "MARCU" }, { "Luca", "LUCA" }, { "Lc.", "LUCA" }, { "Ioan", "IOAN" }, { "Ev. după Ioan", "IOAN" },
            { "Fap.", "FAPTELE APOSTOLILOR" }, { "Romani", "ROMANI" }, { "1 Corinteni", "1 CORINTENI" },
            { "2 Corinteni", "2 CORINTENI" }, { "Galateni", "GALATENI" }, { "Efeseni", "EFESENI" }, { "Ef.", "EFESENI" }, { "Filipeni", "FILIPENI" },
            { "Coloseni", "COLOSENI" }, { "1 Tesaloniceni", "1 TESALONICENI" }, { "2 Tesaloniceni", "2 TESALONICENI" },
            { "1 Timotei", "1 TIMOTEI" }, { "2 Timotei", "2 TIMOTEI" }, { "Tit", "TIT" }, { "Filimon", "FILIMON" },
            { "Evr.", "EVREI" }, { "Iacov", "IACOV" }, { "Iac.", "IACOV" }, { "1 Petru", "1 PETRU" },
            { "1 Pet.", "1 PETRU" }, { "2 Petru", "2 PETRU" }, { "2 Pet.", "2 PETRU" },
            { "1 Ioan", "1 IOAN" }, { "2 Ioan", "2 IOAN" }, { "3 Ioan", "3 IOAN" }, { "Iuda", "IUDA" }, { "Apocalipsa", "APOCALIPSA" }
        };

        public ExplicatiiWindow()
        {
            InitializeComponent();
            comboCarte.ItemsSource = GetCarti();        // din baza de date
            comboExplicator.ItemsSource = GetExplicatori();
        }

      

        private List<string> GetCarti()
        {
            var carti = new List<string>();
            using var conn = new SqliteConnection($"Data Source={CONNECTION_STRING}");
            conn.Open();

            // Varianta corectă: ordonăm după CarteOrdine prin subquery
            using var cmd = new SqliteCommand(@"
        SELECT Carte
        FROM (
            SELECT Carte, MIN(CarteOrdine) AS Ord
            FROM Versete
            GROUP BY Carte
        )
        ORDER BY Ord;", conn);

            using var reader = cmd.ExecuteReader();
            while (reader.Read())
                carti.Add(reader.GetString(0));

            return carti;
        }



        private List<string> GetExplicatori()
        {
            var lista = new List<string>();
            using var conn = new SqliteConnection($"Data Source={CONNECTION_STRING}");
            conn.Open();
            using var cmd = new SqliteCommand("SELECT Nume FROM Explicatori", conn);
            using var reader = cmd.ExecuteReader();
            while (reader.Read()) lista.Add(reader.GetString(0));
            return lista;
        }

        private void AfiseazaExplicatiaPeVerset(object sender, RoutedEventArgs e)
        {
            if (comboCarte.SelectedItem == null ||
                comboCapitol.SelectedItem == null ||
                comboVerset.SelectedItem == null ||
                comboExplicator.SelectedItem == null)
            {
                MessageBox.Show("Te rugăm să selectezi toate câmpurile: carte, capitol, verset și explicator.",
                                "Selecție incompletă",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                return;
            }

            string? carte = (comboCarte.SelectedItem as string)?.Trim();
            int capitol = Convert.ToInt32(comboCapitol.SelectedItem);
            int verset = Convert.ToInt32(comboVerset.SelectedItem);
            string? explicator = (comboExplicator.SelectedItem as string)?.Trim();

            // definim carteCanonica înainte să o folosim
            string carteCanonica = mapareCanonica.ContainsKey(carte) ? mapareCanonica[carte] : carte.ToUpper();

            // titlul de sus
            labelReferinta.Content = $"Explicație pentru {carte} {capitol}:{verset} ({explicator})";

            // versetul original, afișat sub label
            string versetText = GetVersetDinBaza(carteCanonica, capitol, verset);
            versetText = WebUtility.HtmlDecode(versetText);

            textVersetExplicat.Inlines.Clear();
            foreach (var inline in ProceseazaRosu($"{verset}. {versetText}"))
                textVersetExplicat.Inlines.Add(inline);

            // explicația versetului
            string explicatieOriginala = GetExplicatieVerset(carteCanonica, capitol, verset, explicator);
            explicatieOriginala = WebUtility.HtmlDecode(explicatieOriginala);



            // decodăm entitățile HTML (&lt;ROSU&gt; -> <ROSU>)
            explicatieOriginala = WebUtility.HtmlDecode(explicatieOriginala);

            textExplicatieVerset.Inlines.Clear();

            // Regex pentru referințe biblice
            string pattern = @"\b((?:\d\s*)?[A-ZȘȚĂÎa-zșțăî\.]+)\s+(\d+):\s*(\d+(?:\s*[-–]\s*\d+)?(?:\s*,\s*\d+)*)\b";
            var matches = Regex.Matches(explicatieOriginala, pattern);
            int lastIndex = 0;

            foreach (Match match in matches)
            {
                // text normal dintre referințe
                if (match.Index > lastIndex)
                {
                    string intermediar = explicatieOriginala.Substring(lastIndex, match.Index - lastIndex);
                    foreach (var inline in ProceseazaRosu(intermediar))
                        textExplicatieVerset.Inlines.Add(inline);
                }

                // referința biblică
                string carteRef = match.Groups[1].Value.Trim();
                int capitolRef = int.Parse(match.Groups[2].Value);
                string verseteRaw = match.Groups[3].Value;

                string carteRefCanonica = mapareCanonica.ContainsKey(carteRef) ? mapareCanonica[carteRef] : carteRef.ToUpper();

                var segmente = Regex.Split(verseteRaw, @"\s*,\s*");
                foreach (var segment in segmente)
                {
                    string part = Regex.Replace(segment.Trim(), @"[^\d\-–]", "");
                    if (string.IsNullOrEmpty(part)) continue;

                    if (part.Contains("-") || part.Contains("–"))
                    {
                        var parts = Regex.Split(part, @"\s*[-–]\s*");
                        if (parts.Length == 2 && int.TryParse(parts[0], out int start) && int.TryParse(parts[1], out int end))
                        {
                            var versete = GetVerseteDinInterval(carteRefCanonica, capitolRef, start, end);
                            AdaugaReferintaCuPopup(textExplicatieVerset.Inlines, carteRefCanonica, capitolRef, start, end, versete);
                        }
                    }
                    else if (int.TryParse(part, out int versetNr))
                    {
                        var versete = GetVerseteDinInterval(carteRefCanonica, capitolRef, versetNr, versetNr);
                        AdaugaReferintaCuPopup(textExplicatieVerset.Inlines, carteRefCanonica, capitolRef, versetNr, versetNr, versete);
                    }
                }

                lastIndex = match.Index + match.Length;
            }

            // textul rămas după ultima referință
            if (lastIndex < explicatieOriginala.Length)
            {
                string final = explicatieOriginala.Substring(lastIndex);
                foreach (var inline in ProceseazaRosu(final))
                    textExplicatieVerset.Inlines.Add(inline);
            }
        }


        private IEnumerable<Inline> ProceseazaRosu(string text)
        {
            if (string.IsNullOrEmpty(text))
                yield break;

            // caută segmente <ROSU>...</ROSU>
            var regex = new Regex(@"<ROSU>(.*?)</ROSU>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            int lastIndex = 0;

            foreach (Match match in regex.Matches(text))
            {
                if (match.Index > lastIndex)
                {
                    string normal = text.Substring(lastIndex, match.Index - lastIndex);
                    if (!string.IsNullOrEmpty(normal))
                        yield return new Run(normal);
                }

                string rosu = match.Groups[1].Value;
                yield return new Run(rosu)
                {
                    Foreground = Brushes.DarkRed,
                    FontWeight = FontWeights.Bold
                };

                lastIndex = match.Index + match.Length;
            }

            if (lastIndex < text.Length)
            {
                string normal = text.Substring(lastIndex);
                if (!string.IsNullOrEmpty(normal))
                    yield return new Run(normal);
            }
        }







        private void AdaugaReferintaCuPopup(InlineCollection inlines, string carte, int capitol, int start, int end, List<(int Nr, string Text)> versete)
        {
            var run = new Run($"{carte} {capitol}:{start}{(end > start ? "-" + end : "")}")
            {
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.DarkBlue
            };

            var stack = new StackPanel();
            stack.Children.Add(new TextBlock
            {
                Text = $"{carte} {capitol}:{start}{(end > start ? "-" + end : "")}",
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.DarkBlue,
                FontSize = 14,
                Margin = new Thickness(0, 0, 0, 6)
            });

            foreach (var v in versete)
            {
                var tb = new TextBlock
                {
                    TextWrapping = TextWrapping.Wrap,
                    TextAlignment = TextAlignment.Justify,
                    FontSize = 15,
                    FontFamily = new FontFamily("Segoe UI"),
                    FontWeight = FontWeights.SemiBold,
                    Foreground = new SolidColorBrush(Color.FromRgb(30, 30, 30)),
                    LineHeight = 20,
                    MaxWidth = 420,
                    Margin = new Thickness(0, 2, 0, 6),
                    Padding = new Thickness(4, 0, 4, 0)
                };

                // Procesăm tagurile <ROSU> din verset
                string versetText = WebUtility.HtmlDecode($"{v.Nr}. {v.Text}");
                foreach (var inline in ProceseazaRosu(versetText))
                    tb.Inlines.Add(inline);

                stack.Children.Add(tb);
            }

            var popup = CreeazaPopup(stack);
            run.MouseLeftButtonDown += (s, ev) =>
            {
                popup.IsOpen = true;
                ev.Handled = true;
            };

            this.PreviewMouseDown += (s, ev) =>
            {
                if (popup.IsOpen)
                    popup.IsOpen = false;
            };

            inlines.Add(run);
            inlines.Add(new Run(" "));
        }


        private Popup CreeazaPopup(StackPanel stack)
        {
            var scroll = new ScrollViewer
            {
                Content = stack,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                MaxHeight = 300,
                Focusable = true
            };

            return new Popup
            {
                Placement = PlacementMode.Mouse,
                StaysOpen = true,
                Child = new Border
                {
                    Background = Brushes.LightYellow,
                    BorderBrush = Brushes.DarkGoldenrod,
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(6),
                    Padding = new Thickness(8),
                    Margin = new Thickness(2),
                    Effect = new System.Windows.Media.Effects.DropShadowEffect
                    {
                        Color = Colors.OrangeRed,
                        BlurRadius = 12,
                        ShadowDepth = 2,
                        Opacity = 0.4
                    },
                    Child = scroll
                }
            };
        }





        // Helper: returnează lista de versete din interval
        private List<(int Nr, string Text)> GetVerseteDinInterval(string carte, int capitol, int start, int end)
        {
            var lista = new List<(int, string)>();
            using var conn = new SqliteConnection($"Data Source={CONNECTION_STRING}");
            conn.Open();

            using var cmd = new SqliteCommand(@"
        SELECT Verset, Text 
        FROM Versete 
        WHERE UPPER(Carte) = @carte 
        AND Capitol = @capitol 
        AND Verset BETWEEN @start AND @end 
        ORDER BY Verset", conn);

            cmd.Parameters.AddWithValue("@carte", carte);
            cmd.Parameters.AddWithValue("@capitol", capitol);
            cmd.Parameters.AddWithValue("@start", start);
            cmd.Parameters.AddWithValue("@end", end);

            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                int nr = reader.GetInt32(0);
                string text = reader.GetString(1);
                lista.Add((nr, text));
            }

            return lista;
        }

        private string GetExplicatieVerset(string carte, int capitol, int verset, string explicator)
        {
            carte = carte.Trim().ToUpper();
            explicator = explicator.Trim().ToUpper();

            using var conn = new SqliteConnection($"Data Source={CONNECTION_STRING}");
            conn.Open();

            using var cmd = new SqliteCommand(@"
        SELECT e.TextExplicativ
        FROM ExplicatiiVerset e
        JOIN Explicatori x ON e.IdExplicator = x.Id
        WHERE UPPER(e.Carte) = @carte
        AND e.Capitol = @capitol
        AND e.Verset = @verset
        AND UPPER(x.Nume) = @explicator", conn);

            cmd.Parameters.AddWithValue("@carte", carte);
            cmd.Parameters.AddWithValue("@capitol", capitol);
            cmd.Parameters.AddWithValue("@verset", verset);
            cmd.Parameters.AddWithValue("@explicator", explicator);

            var result = cmd.ExecuteScalar();
            return result?.ToString() ?? "Nu există explicație de la acest autor pentru acest verset.";
        }

        private string GetVersetDinBaza(string carte, int capitol, int verset)
        {
            carte = carte.Trim().ToUpper();

            using var conn = new SqliteConnection($"Data Source={CONNECTION_STRING}");
            conn.Open();

            using var cmd = new SqliteCommand(@"
        SELECT Text FROM Versete
        WHERE UPPER(Carte) = @carte
        AND Capitol = @capitol
        AND Verset = @verset", conn);

            cmd.Parameters.AddWithValue("@carte", carte);
            cmd.Parameters.AddWithValue("@capitol", capitol);
            cmd.Parameters.AddWithValue("@verset", verset);

            var result = cmd.ExecuteScalar();
            return result?.ToString() ?? $"[{carte} {capitol}:{verset}] nu există în baza.";
        }

        private List<int> GetCapitole(string carte)
        {
            var lista = new List<int>();
            using var conn = new SqliteConnection($"Data Source={CONNECTION_STRING}");
            conn.Open();
            using var cmd = new SqliteCommand(@"
        SELECT DISTINCT Capitol 
        FROM Versete 
        WHERE Carte = @carte 
        ORDER BY Capitol", conn);
            cmd.Parameters.AddWithValue("@carte", carte.ToUpper());

            using var reader = cmd.ExecuteReader();
            while (reader.Read()) lista.Add(reader.GetInt32(0));
            return lista;
        }

        private List<int> GetVersete(string carte, int capitol)
        {
            var lista = new List<int>();
            using var conn = new SqliteConnection($"Data Source={CONNECTION_STRING}");
            conn.Open();
            using var cmd = new SqliteCommand(@"
        SELECT Verset 
        FROM Versete 
        WHERE Carte = @carte AND Capitol = @capitol 
        ORDER BY Verset", conn);
            cmd.Parameters.AddWithValue("@carte", carte.ToUpper());
            cmd.Parameters.AddWithValue("@capitol", capitol);

            using var reader = cmd.ExecuteReader();
            while (reader.Read()) lista.Add(reader.GetInt32(0));
            return lista;
        }
        private void comboCarte_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (comboCarte.SelectedItem is string carte)
            {
                string canonica = mapareCanonica.ContainsKey(carte) ? mapareCanonica[carte] : carte.ToUpper();
                comboCapitol.ItemsSource = GetCapitole(canonica);
                comboVerset.ItemsSource = null; // golește versetele până alegi capitolul
            }
        }

        private void comboCapitol_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (comboCarte.SelectedItem is string carte && comboCapitol.SelectedItem is int capitol)
            {
                string canonica = mapareCanonica.ContainsKey(carte) ? mapareCanonica[carte] : carte.ToUpper();
                comboVerset.ItemsSource = GetVersete(canonica, capitol);
            }
        }

        private void CapitolUp_Click(object sender, RoutedEventArgs e)
        {
            if (comboCapitol.SelectedItem is int current)
            {
                int index = comboCapitol.Items.IndexOf(current);
                if (index < comboCapitol.Items.Count - 1)
                    comboCapitol.SelectedIndex = index + 1;
            }

            AfiseazaExplicatiaPeVerset(null, null); // actualizează tot
        }


        private void CapitolDown_Click(object sender, RoutedEventArgs e)
        {
            if (comboCapitol.SelectedItem is int current)
            {
                int index = comboCapitol.Items.IndexOf(current);
                if (index > 0)
                    comboCapitol.SelectedIndex = index - 1;
            }

            AfiseazaExplicatiaPeVerset(null, null);
        }


        private void VersetUp_Click(object sender, RoutedEventArgs e)
        {
            if (comboVerset.SelectedItem is int current)
            {
                int index = comboVerset.Items.IndexOf(current);
                if (index < comboVerset.Items.Count - 1)
                    comboVerset.SelectedIndex = index + 1;
            }

            AfiseazaExplicatiaPeVerset(null, null);
        }


        private void VersetDown_Click(object sender, RoutedEventArgs e)
        {
            if (comboVerset.SelectedItem is int current)
            {
                int index = comboVerset.Items.IndexOf(current);
                if (index > 0)
                    comboVerset.SelectedIndex = index - 1;
            }

            AfiseazaExplicatiaPeVerset(null, null);
        }



    }
}