using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Run = System.Windows.Documents.Run;

namespace Biblia
{
    public partial class FereastraFavoriteView : Window
    {
        private readonly Dictionary<string, List<VersetSchita>> _favoritePeListe;
        
        private readonly List<VersetSchita> _schitaDePredica = new();

        public FereastraFavoriteView(Dictionary<string, List<VersetSchita>> favoritePeListe)
        {
            InitializeComponent();

            _favoritePeListe = favoritePeListe ?? new();

            comboListe.ItemsSource = _favoritePeListe.Keys.ToList();
            if (comboListe.Items.Count > 0)
                comboListe.SelectedIndex = 0;
        }

        private void comboListe_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            listBoxFavorite.Items.Clear();

            if (comboListe.SelectedItem is string lista &&
                _favoritePeListe.TryGetValue(lista, out var versete))
            {
                foreach (var v in versete)
                {
                    var tb = new TextBlock
                    {
                        TextWrapping = TextWrapping.Wrap,
                        FontSize = 14
                    };

                    tb.Inlines.Add(new Run($"{v.Carte} {v.Capitol}:{v.Verset} — ")
                    { FontWeight = FontWeights.Bold });

                    foreach (var inline in ParseVersetCuRosu(v.Text))
                        tb.Inlines.Add(inline);

                    if (!string.IsNullOrWhiteSpace(v.Comentariu))
                    {
                        tb.Inlines.Add(new LineBreak());
                        tb.Inlines.Add(new Run(v.Comentariu)
                        {
                            Foreground = Brushes.DarkBlue,
                            FontStyle = FontStyles.Italic,
                            FontSize = 12
                        });
                    }

                    listBoxFavorite.Items.Add(new ListBoxItem { Content = tb });
                }
            }
        }

        private List<Inline> ParseVersetCuRosu(string text)
        {
            var inlines = new List<Inline>();
            int index = 0;

            while (index < text.Length)
            {
                int start = text.IndexOf("<ROSU>", index);
                if (start == -1)
                {
                    inlines.Add(new Run(text.Substring(index)));
                    break;
                }

                if (start > index)
                    inlines.Add(new Run(text.Substring(index, start - index)));

                int end = text.IndexOf("</ROSU>", start);
                if (end == -1)
                {
                    inlines.Add(new Run(text.Substring(start)) { Foreground = Brushes.Red });
                    break;
                }

                string rosuText = text.Substring(start + 6, end - start - 6);
                inlines.Add(new Run(rosuText) { Foreground = Brushes.Red });

                index = end + 7;
            }

            return inlines;
        }

        // 🗑️ Ștergere verset
        private void BtnSterge_Click(object sender, RoutedEventArgs e)
        {
            if (listBoxFavorite.SelectedItem is ListBoxItem item &&
                item.Content is TextBlock tb)
            {
                string lista = comboListe.SelectedItem as string;
                if (string.IsNullOrEmpty(lista))
                    return;

                // găsim versetul selectat
                var verset = _favoritePeListe[lista].FirstOrDefault(v =>
                    tb.Inlines.FirstInline is Run r &&
                    r.Text.StartsWith($"{v.Carte} {v.Capitol}:{v.Verset}"));

                if (verset != null)
                {
                    // 🔥 apelăm metoda din BibliaWindow
                    if (Owner is BibliaWindow main)
                    {
                        main.StergeVersetDinLista(lista, verset);

                        // 🔥 reîncărcăm structura în memorie
                        _favoritePeListe.Clear();
                        foreach (var kv in main.IncarcaFavoriteDinBazaDeDate())
                            _favoritePeListe[kv.Key] = kv.Value;
                    }

                    // 🔥 reafișăm lista
                    comboListe_SelectionChanged(null, null);
                }
            }
        }


        // ✏️ Editare comentariu
        private void BtnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (listBoxFavorite.SelectedItem is ListBoxItem item &&
                item.Content is TextBlock tb)
            {
                string lista = comboListe.SelectedItem as string;
                if (string.IsNullOrEmpty(lista)) return;

                var verset = _favoritePeListe[lista].FirstOrDefault(v =>
                    tb.Inlines.FirstInline is Run r &&
                    r.Text.StartsWith($"{v.Carte} {v.Capitol}:{v.Verset}"));

                if (verset != null)
                {
                    var f = new FereastraEditareComentariu(verset.Comentariu);

                    if (f.ShowDialog() == true)
                    {
                        verset.Comentariu = f.ComentariuFinal;

                        if (Owner is BibliaWindow main)
                        {
                            // 🔥 salvăm în DB
                            main.ActualizeazaComentariu(lista, verset);

                            // 🔥 reîncărcăm structura din DB
                            _favoritePeListe.Clear();
                            foreach (var kv in main.IncarcaFavoriteDinBazaDeDate())
                                _favoritePeListe[kv.Key] = kv.Value;
                        }

                        // 🔥 reafișăm lista
                        comboListe_SelectionChanged(null, null);

                    }
                }
            }
        }


        // ➕ Trimite la schiță
        private void BtnTrimiteLaSchita_Click(object sender, RoutedEventArgs e)
        {
            if (listBoxFavorite.SelectedItem is ListBoxItem item &&
                item.Content is TextBlock tb)
            {
                string lista = comboListe.SelectedItem as string;
                if (string.IsNullOrEmpty(lista)) return;

                var verset = _favoritePeListe[lista].FirstOrDefault(v =>
                    tb.Inlines.FirstInline is Run r &&
                    r.Text.StartsWith($"{v.Carte} {v.Capitol}:{v.Verset}"));

                if (verset != null && Owner is BibliaWindow main)
                {
                    main.AdaugaLaSchita(verset);

                    if (main.EsteModSchitaActiv &&
                        main.CarteCurenta == verset.Carte &&
                        main.CapitolCurent == verset.Capitol)
                    {
                        main.IncarcaVersete(verset.Carte, verset.Capitol);
                    }
                }
            }
        }

        // 🗑️ Ștergere listă
        private void BtnStergeLista_Click(object sender, RoutedEventArgs e)
        {
            if (comboListe.SelectedItem is string lista)
            {
                var confirm = MessageBox.Show(
                    $"Sigur vrei să ștergi lista '{lista}'?",
                    "Confirmare ștergere listă",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (confirm == MessageBoxResult.Yes)
                {
                    // 🔥 apelăm metoda din BibliaWindow
                    if (Owner is BibliaWindow main)
                    {
                        main.StergeLista(lista);

                        // 🔥 reîncărcăm favoritele din DB
                        _favoritePeListe.Clear();
                        foreach (var kv in main.IncarcaFavoriteDinBazaDeDate())
                            _favoritePeListe[kv.Key] = kv.Value;
                    }

                    // 🔥 actualizăm UI-ul
                    comboListe.ItemsSource = _favoritePeListe.Keys.ToList();

                    if (comboListe.Items.Count > 0)
                        comboListe.SelectedIndex = 0;
                    else
                        listBoxFavorite.Items.Clear();
                }
            }
        }


        // ✏️ Redenumire listă
        private void BtnEditLista_Click(object sender, RoutedEventArgs e)
        {
            if (comboListe.SelectedItem is string lista)
            {
                var input = Microsoft.VisualBasic.Interaction.InputBox(
                    "Scrie noul nume pentru listă:",
                    "Redenumește listă",
                    lista);

                if (!string.IsNullOrWhiteSpace(input))
                {
                    string nouNume = input.Trim();

                    if (!_favoritePeListe.ContainsKey(nouNume))
                    {
                        // 🔥 apelăm metoda din BibliaWindow
                        if (Owner is BibliaWindow main)
                        {
                            main.RedenumesteLista(lista, nouNume);

                            // 🔥 reîncărcăm favoritele din DB
                            _favoritePeListe.Clear();
                            foreach (var kv in main.IncarcaFavoriteDinBazaDeDate())
                                _favoritePeListe[kv.Key] = kv.Value;
                        }

                        // 🔥 actualizăm UI-ul
                        comboListe.ItemsSource = _favoritePeListe.Keys.ToList();
                        comboListe.SelectedItem = nouNume;
                    }
                    else
                    {
                        MessageBox.Show("Există deja o listă cu acest nume.");
                    }
                }
            }
        }


        // 💾 Butonul de salvare (sus și jos)
        private void BtnSalveaza_Click(object sender, RoutedEventArgs e)
        {
            
            MessageBox.Show("Favoritele au fost salvate.");
        }


    }
}
