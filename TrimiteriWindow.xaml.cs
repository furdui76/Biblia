using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace Biblia
{
    public partial class TrimiteriWindow : Window
    {

        public void LoadTema(TemaItem tema)
        {
            titluTema.Text = $"Versete pentru: {tema.Nume}";
            comboCarte.SelectedIndex = -1;
            comboCapitol.SelectedIndex = -1;
            listVersete.Items.Clear();

            AfiseazaVersetePeTema(tema.Id, tema.Nume);
            AfiseazaBlocuriPeTema(tema.Nume);
        }
        public TrimiteriWindow(TemaItem tema)
        {
            InitializeComponent();
            titluTema.Text = $"Versete pentru: {tema.Nume}";
            comboCarte.ItemsSource = GetCarti();
            comboCapitol.ItemsSource = Enumerable.Range(1, 150).ToList();
            AfiseazaVersetePeTema(tema.Id, tema.Nume);
            AfiseazaBlocuriPeTema(tema.Nume);


        }
       


        private void AfiseazaVerseteCuTrimiteri(string tema)
        {
            string[] keywords = GetKeywords(tema);
            using var conn = new SQLiteConnection("Data Source=biblia.db");
            conn.Open();

            var conditions = string.Join(" OR ", keywords.Select((kw, i) => $"Text LIKE '%' || @kw{i} || '%' COLLATE NOCASE"));
            var cmd = new SQLiteCommand($@"
        SELECT IdVerset, Carte, Capitol, Verset, Text
        FROM Versete
        WHERE ({conditions})
        ORDER BY CarteOrdine, Capitol, Verset;", conn);

            for (int i = 0; i < keywords.Length; i++)
                cmd.Parameters.AddWithValue($"@kw{i}", keywords[i]);

            using var reader = cmd.ExecuteReader();
            listVersete.Items.Clear();
            int count = 0;

            while (reader.Read())
            {
                int versetId = reader.GetInt32(0);
                string carte = reader.GetString(1);
                int capitol = reader.GetInt32(2);
                int verset = reader.GetInt32(3);
                string text = reader.GetString(4);

                string full = $"{carte} {capitol}:{verset} — {text}";
                var tb = new TextBlock { TextWrapping = TextWrapping.Wrap, FontSize = 16 };

                var parts = HighlightKeywords(full, keywords);
                foreach (var part in parts) tb.Inlines.Add(part);

                listVersete.Items.Add(tb);
                count++; // ✅ Numărăm DOAR versetele
            }

            // Separat: trimiterile
            var cmdTrim = new SQLiteCommand("SELECT Tip, Referinta FROM TReferinta WHERE VersetId IN (SELECT IdVerset FROM Versete WHERE " + conditions + ")", conn);
            for (int i = 0; i < keywords.Length; i++)
                cmdTrim.Parameters.AddWithValue($"@kw{i}", keywords[i]);

            using var readerTrim = cmdTrim.ExecuteReader();
            while (readerTrim.Read())
            {
                string tip = readerTrim.GetString(0);
                string referinta = readerTrim.GetString(1);

                var tbTrim = new TextBlock
                {
                    FontSize = 14,
                    TextWrapping = TextWrapping.Wrap
                };
                tbTrim.Inlines.Add(new Run("↳ " + tip + ": "));
                tbTrim.Inlines.Add(new Run(referinta) { Foreground = Brushes.DarkSlateBlue });

                var border = new Border
                {
                    BorderBrush = Brushes.SlateGray,
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(6),
                    Margin = new Thickness(20, 0, 0, 8),
                    Padding = new Thickness(8),
                    Background = Brushes.WhiteSmoke,
                    Child = tbTrim
                };

                listVersete.Items.Add(border);
            }

            labelCount.Content = count == 0
                ? "⚠️ Niciun verset găsit."
                : $"🔍 {count} verset{(count == 1 ? "" : "e")} găsit{(count == 1 ? "" : "e")} pentru tema {tema}.";
        }





        private void OnFiltreazaClick(object sender, RoutedEventArgs e)
        {
            string carte = comboCarte.SelectedItem as string;
            int capitol = comboCapitol.SelectedItem is int val ? val : -1;
            string[] keywords = GetKeywords(titluTema.Text);

            using var conn = new SQLiteConnection("Data Source=biblia.db");
            conn.Open();

            var conditions = string.Join(" OR ", keywords.Select((kw, i) => $"Text LIKE '%' || @kw{i} || '%' COLLATE NOCASE"));
            string filter = "";
            if (!string.IsNullOrEmpty(carte)) filter += " AND Carte = @carte";
            if (capitol > 0) filter += " AND Capitol = @capitol";

            var cmd = new SQLiteCommand($@"
        SELECT Carte, Capitol, Verset, Text
        FROM Versete
        WHERE ({conditions}) {filter}
        ORDER BY CarteOrdine, Capitol, Verset;", conn);

            for (int i = 0; i < keywords.Length; i++)
                cmd.Parameters.AddWithValue($"@kw{i}", keywords[i]);
            if (!string.IsNullOrEmpty(carte)) cmd.Parameters.AddWithValue("@carte", carte);
            if (capitol > 0) cmd.Parameters.AddWithValue("@capitol", capitol);

            using var reader = cmd.ExecuteReader();
            listVersete.Items.Clear();
            int count = 0;

            while (reader.Read())
            {
                string carteRez = reader.GetString(0);
                int capitolRez = reader.GetInt32(1);
                int verset = reader.GetInt32(2);
                string text = reader.GetString(3);

                var tb = new TextBlock { TextWrapping = TextWrapping.Wrap, FontSize = 16 };
                string full = $"{carteRez} {capitolRez}:{verset} — {text}";
                var parts = HighlightKeywords(full, keywords);
                foreach (var part in parts) tb.Inlines.Add(part);
                listVersete.Items.Add(tb);
                count++;
            }

            labelCount.Content = count == 0
                ? "⚠️ Niciun verset găsit."
                : $"🔍 {count} verset{(count == 1 ? "" : "e")} găsit{(count == 1 ? "" : "e")}.";
        }


        private void OnReseteazaClick(object sender, RoutedEventArgs e)
        {
            comboCarte.SelectedIndex = -1;
            comboCapitol.SelectedIndex = -1;
            AfiseazaVerseteCuTrimiteri(titluTema.Text);
        }

        private string[] GetKeywords(string tema)
        {
            if (tema.Contains("adevăr", StringComparison.OrdinalIgnoreCase))
                return new[] { "Adevăr", "adevărul", "adevărat", "adevărată", "adevărați", "adevăratul" };

            if (tema.Contains("ascult", StringComparison.OrdinalIgnoreCase))
                return new[] { "Ascultare", "ascult", "ascultă", "ascultând", "ascultător", "ascultarea" };

            if (tema.Contains("botez", StringComparison.OrdinalIgnoreCase))
                return new[] { "Botez", "botezul", "botezați", "botezând", "s-a botezat", "m-am botezat", "botez în apă" };

            if (tema.Contains("credin", StringComparison.OrdinalIgnoreCase))
                return new[] { "Credința", "credinţă", "crede", "crezut", "credincios", "încredinţat" };

            if (tema.Contains("dragost", StringComparison.OrdinalIgnoreCase) || tema.Contains("iubir", StringComparison.OrdinalIgnoreCase))
                return new[] { "Dragoste", "iubire", "iubește", "iubit", "iubind", "iubirea", "dragostea" };

            if (tema.Contains("duhul", StringComparison.OrdinalIgnoreCase))
                return new[] { "Duhul Sfânt", "Duhul", "Sfântul Duh", "Duhului", "Duhul lui Dumnezeu", "Duhul adevărului" };

            if (tema.Contains("har", StringComparison.OrdinalIgnoreCase))
                return new[] { "Har", "harul", "haruri", "prin har", "harul lui Dumnezeu" };

            if (tema.Contains("iertar", StringComparison.OrdinalIgnoreCase))
                return new[] { "Iertare", "iartă", "iertat", "ne iartă", "iertând", "iartă-ne", "iertându-i" };

            if (tema.Contains("înțelep", StringComparison.OrdinalIgnoreCase))
                return new[] { "Înțelepciune", "înțelept", "înțelepți", "înțelepciunea", "înțelepțește", "pricepere" };

            if (tema.Contains("împărăția", StringComparison.OrdinalIgnoreCase))
                return new[] { "Împărăția", "împărăţia", "împărat", "domnie", "tron", "împărăției", "împărăția lui Dumnezeu" };

            if (tema.Contains("jertf", StringComparison.OrdinalIgnoreCase))
                return new[] { "Jertfă", "jertfa", "jertfit", "jertfește", "jertfind", "jertfe", "jertfa Lui" };

            if (tema.Contains("judec", StringComparison.OrdinalIgnoreCase))
                return new[] { "Judecată", "judecata", "judecă", "judecând", "judecător", "judecății", "judecat" };

            if (tema.Contains("lumin", StringComparison.OrdinalIgnoreCase))
                return new[] { "Lumină", "lumina", "luminează", "luminează-ne", "luminos", "luminii", "luminează-mă" };

            if (tema.Contains("mesia", StringComparison.OrdinalIgnoreCase))
                return new[] { "Mesia", "unsul", "hristos", "mântuitor" };

            if (tema.Contains("milă", StringComparison.OrdinalIgnoreCase))
                return new[] { "Milă", "mila", "milostiv", "milostivire", "îndurare", "îndurarea" };

            if (tema.Contains("mântuir", StringComparison.OrdinalIgnoreCase))
                return new[] { "Mântuire", "mântuirea", "mântuit", "mântuiește", "mântuind", "mântuitor" };

            if (tema.Contains("mulțum", StringComparison.OrdinalIgnoreCase))
                return new[] { "Mulțumire", "mulțumiți", "mulțumesc", "recunoștință", "recunoscători", "mulțumindu-i" };

            if (tema.Contains("nădejde", StringComparison.OrdinalIgnoreCase))
                return new[] { "Nădejde", "nădejdea", "nădăjduiește", "nădăjduit", "nădăjduind" };

            if (tema.Contains("neprihăn", StringComparison.OrdinalIgnoreCase))
                return new[] { "Neprihănire", "neprihănit", "neprihănită", "neprihăniți", "neprihănirea", "neprihănitului" };

            if (tema.Contains("păcat", StringComparison.OrdinalIgnoreCase))
                return new[] { "Păcat", "păcatul", "păcate", "păcatele", "păcătos", "păcătoși", "păcătuiește", "păcătuind" };

            if (tema.Contains("rugăciune", StringComparison.OrdinalIgnoreCase))
                return new[] { "Rugăciune", "roagă", "mă rog", "s-a rugat" };

            if (tema.Contains("sfinț", StringComparison.OrdinalIgnoreCase))
                return new[] { "Sfânt", "sfântă", "sfințenie", "sfințit", "sfințește", "sfințind", "sfinți" };

            if (tema.Contains("veșnic", StringComparison.OrdinalIgnoreCase))
                return new[] { "Viață veșnică", "viața veșnică", "veșnicie", "veșnic", "veșnică", "veșnice", "veșnicia" };

            return new[] { "DUMNEZEU" };
        }

        private void AfiseazaVersetePeTema(int idTema, string numeTema)
{
    using var conn = new SQLiteConnection("Data Source=biblia.db");
    conn.Open();

    // 1. Versete legate explicit de temă
    var cmdTema = new SQLiteCommand(@"
        SELECT IdVerset, Carte, Capitol, Verset, Text
        FROM Versete
        WHERE ',' || Trimiteri || ',' LIKE '%,' || @idTema || ',%'
        ORDER BY CarteOrdine, Capitol, Verset", conn);

    cmdTema.Parameters.AddWithValue("@idTema", idTema);

    using var readerTema = cmdTema.ExecuteReader();
    listVersete.Items.Clear();
    int count = 0;

    while (readerTema.Read())
    {
        int versetId = readerTema.GetInt32(0);
        string carte = readerTema.GetString(1);
        int capitol = readerTema.GetInt32(2);
        int verset = readerTema.GetInt32(3);
        string text = readerTema.GetString(4);

        string full = $"{carte} {capitol}:{verset} — {text}";
        var tb = new TextBlock { TextWrapping = TextWrapping.Wrap, FontSize = 16 };

        // Evidențiere cuvânt-cheie
        var parts = HighlightKeywords(full, GetKeywords(numeTema));
        foreach (var part in parts) tb.Inlines.Add(part);

        listVersete.Items.Add(tb);
        count++;

        // Diagnostic opțional
        Console.WriteLine($"[Trimiteri] VersetId: {versetId}");
    }

    // 2. Fallback dacă nu s-a găsit nimic
    if (count == 0)
    {
        MessageBox.Show("⚠️ Nu există versete legate explicit de această temă. Se caută sugestii...");
        AfiseazaVerseteCuTrimiteri(numeTema);
    }

    labelCount.Content = count == 0
        ? "⚠️ Niciun verset găsit explicit."
        : $"🔍 {count} verset{(count == 1 ? "" : "e")} găsit{(count == 1 ? "" : "e")} pentru tema {numeTema}.";
}

        private void AfiseazaBlocuriPeTema(string tema)
        {
            string[] keywords = GetKeywords(tema);
            using var conn = new SQLiteConnection("Data Source=biblia.db");
            conn.Open();

            listVersete.Items.Clear();
            var blocuriAfisate = new HashSet<string>();
            int count = 0;

            foreach (var kw in keywords)
            {
                var cmd = new SQLiteCommand(@"
            SELECT Carte, Capitol, Verset, Text
            FROM Versete
            WHERE Text LIKE '%' || @kw || '%' COLLATE NOCASE
            ORDER BY CarteOrdine, Capitol, Verset;", conn);

                cmd.Parameters.AddWithValue("@kw", kw);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string carte = reader.GetString(0);
                    int capitol = reader.GetInt32(1);
                    int verset = reader.GetInt32(2);

                    int start = Math.Max(1, verset - 2);
                    int end = verset + 2;
                    string blocId = $"{carte}-{capitol}-{start}-{end}";

                    if (blocuriAfisate.Contains(blocId)) continue;
                    blocuriAfisate.Add(blocId);

                    var cmdBloc = new SQLiteCommand(@"
                SELECT Verset, Text
                FROM Versete
                WHERE Carte = @carte AND Capitol = @capitol
                  AND Verset BETWEEN @start AND @end
                ORDER BY Verset;", conn);

                    cmdBloc.Parameters.AddWithValue("@carte", carte);
                    cmdBloc.Parameters.AddWithValue("@capitol", capitol);
                    cmdBloc.Parameters.AddWithValue("@start", start);
                    cmdBloc.Parameters.AddWithValue("@end", end);

                    using var readerBloc = cmdBloc.ExecuteReader();
                    string blocText = "";
                    var verseteBloc = new List<string>();

                    while (readerBloc.Read())
                    {
                        int v = readerBloc.GetInt32(0);
                        string t = readerBloc.GetString(1);
                        verseteBloc.Add($"{v}. {t}");
                        blocText += t + " ";
                    }

                    // Filtrare pe cuvânt exact
                    bool contineExact = keywords.Any(k =>
                        Regex.IsMatch(blocText, $@"\b{Regex.Escape(k)}\b", RegexOptions.IgnoreCase));
                    if (!contineExact) continue;

                    // Construim TextBlock cu wrap și evidențiere
                    var tb = new TextBlock
                    {
                        TextWrapping = TextWrapping.Wrap,
                        FontSize = 16,
                        MaxWidth = 740,
                        Margin = new Thickness(0, 0, 0, 12)
                    };

                    string finalText = $"{carte} {capitol}:{start}-{end} — {string.Join(" ", verseteBloc)}";

                    // Dacă există taguri <ROSU>, parsează-le
                    if (finalText.Contains("<ROSU>"))
                    {
                        var tbRosu = ParseRosuTags(finalText);
                        listVersete.Items.Add(tbRosu);
                    }
                    else
                    {
                        var parts = HighlightKeywords(finalText, keywords);
                        foreach (var part in parts) tb.Inlines.Add(part);
                        listVersete.Items.Add(tb);
                    }

                    count++;
                }
            }

            labelCount.Content = count == 0
                ? $"⚠️ Niciun bloc găsit pentru tema {tema}."
                : $"📚 {count} bloc{(count == 1 ? "" : "uri")} găsit{(count == 1 ? "" : "e")} pentru tema {tema}.";
        }




        private Inline[] HighlightKeywords(string text, string[] keywords)
        {
            var inlines = new List<Inline>();
            int index = 0;
            string lower = text.ToLower();

            var matches = keywords
    .SelectMany(kw => AllIndexesOf(lower, kw.ToLower())
        .Select(pos => (pos, kw)))
    .OrderBy(m => m.pos)
    .ToList();

            // Elimină suprapunerile
            var distinctMatches = new List<(int pos, string kw)>();
            int lastEnd = -1;
            foreach (var match in matches)
            {
                if (match.pos >= lastEnd)
                {
                    distinctMatches.Add(match);
                    lastEnd = match.pos + match.kw.Length;
                }
            }


            foreach (var match in distinctMatches)

            {
                if (match.pos > index)
                {
                    string normal = text.Substring(index, match.pos - index);
                    inlines.Add(new Run(normal));
                }

                // Folosește cuvântul-cheie direct, nu din textul original
                inlines.Add(new Run(match.kw)
                {
                    Foreground = Brushes.DarkRed,
                    FontWeight = FontWeights.Bold
                });

                index = match.pos + match.kw.Length;
            }

            if (index < text.Length)
                inlines.Add(new Run(text.Substring(index)));

            return inlines.ToArray();
        }


        private List<int> AllIndexesOf(string text, string keyword)
        {
            var indexes = new List<int>();
            int index = 0;
            while ((index = text.IndexOf(keyword, index, StringComparison.OrdinalIgnoreCase)) != -1)
            {
                indexes.Add(index);
                index += keyword.Length;
            }
            return indexes;
        }

        private List<string> GetCarti()
        {
            using var conn = new SQLiteConnection("Data Source=biblia.db");
            conn.Open();

            var cmd = new SQLiteCommand("SELECT DISTINCT Carte FROM Versete ORDER BY Carte;", conn);
            using var reader = cmd.ExecuteReader();

            var carti = new List<string>();
            while (reader.Read())
                carti.Add(reader.GetString(0));
            return carti;
        }

        private TextBlock ParseRosuTags(string input)
        {
            var tb = new TextBlock
            {
                TextWrapping = TextWrapping.Wrap,
                FontSize = 16,
                MaxWidth = 740,
                Margin = new Thickness(0, 0, 0, 12)
            };

            int pos = 0;
            while (pos < input.Length)
            {
                int startTag = input.IndexOf("<ROSU>", pos, StringComparison.OrdinalIgnoreCase);
                if (startTag == -1)
                {
                    tb.Inlines.Add(new Run(input.Substring(pos)));
                    break;
                }

                if (startTag > pos)
                    tb.Inlines.Add(new Run(input.Substring(pos, startTag - pos)));

                int endTag = input.IndexOf("</ROSU>", startTag, StringComparison.OrdinalIgnoreCase);
                if (endTag == -1)
                {
                    tb.Inlines.Add(new Run(input.Substring(startTag)) { Foreground = Brushes.Red, FontWeight = FontWeights.Bold });
                    break;
                }

                string rosuText = input.Substring(startTag + 6, endTag - startTag - 6);
                tb.Inlines.Add(new Run(rosuText) { Foreground = Brushes.Red, FontWeight = FontWeights.Bold });

                pos = endTag + 7;
            }

            return tb;
        }
       
       

       


    }
}








