using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;


namespace Biblia
{
    public partial class FereastraComentariu : Window
    {
        private SQLiteConnection conn;
        private string carte;
        private int capitol;
        private int verset;
        private Action<string> adaugaLaSchitaCallback;
        private FlowDocument flowDoc;


        public FereastraComentariu(SQLiteConnection conn, string carte, int capitol, int verset)
        {
            InitializeComponent();
            richComentariu.IsDocumentEnabled = true;   // permite Hyperlink.Click
            richComentariu.IsReadOnly = true;          // opțional: previne editarea

            

            this.conn = conn;
            this.carte = carte;
            this.capitol = capitol;
            this.verset = verset;

            flowDoc = new FlowDocument();
            richComentariu.Document = flowDoc;

            IncarcaAutori();
        }



        private void IncarcaAutori()
        {
            var cmd = new SQLiteCommand("SELECT AutorComentariu FROM Comentarii WHERE Carte = @carte AND Capitol = @capitol AND Verset = @verset", conn);
            cmd.Parameters.AddWithValue("@carte", carte);
            cmd.Parameters.AddWithValue("@capitol", capitol);
            cmd.Parameters.AddWithValue("@verset", verset);

            var reader = cmd.ExecuteReader();
            var autori = new List<string>();

            while (reader.Read())
            {
                autori.Add(reader.GetString(0));
            }

            comboAutori.ItemsSource = autori;
            if (autori.Count > 0)
                comboAutori.SelectedIndex = 0;
        }

        private void comboAutori_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (comboAutori.SelectedItem == null) return;

            var autor = comboAutori.SelectedItem.ToString();

            var cmd = new SQLiteCommand(
                "SELECT TextComentariu FROM Comentarii WHERE Carte = @carte AND Capitol = @capitol AND Verset = @verset AND AutorComentariu = @autor",
                conn
            );
            cmd.Parameters.AddWithValue("@carte", carte);
            cmd.Parameters.AddWithValue("@capitol", capitol);
            cmd.Parameters.AddWithValue("@verset", verset);
            cmd.Parameters.AddWithValue("@autor", autor);

            var rezultat = cmd.ExecuteScalar()?.ToString();
            if (string.IsNullOrWhiteSpace(rezultat)) return;

            // 🔍 Extragem indicatorul dintre paranteze
            var match = Regex.Match(rezultat, @"\((.*?)\)");
            string indicatorText = match.Success ? match.Groups[1].Value.Trim() : "";

            // 🧹 Curățăm comentariul (fără indicator)
            string textFaraIndicator = rezultat;
            if (match.Success)
            {
                int indexSfarsit = rezultat.IndexOf(')') + 1;
                if (indexSfarsit < rezultat.Length)
                    textFaraIndicator = rezultat.Substring(indexSfarsit).TrimStart('\n', '\r', ' ');
            }
            AfiseazaComentariu(textFaraIndicator);

            // 🧠 Construim indicatorul afișat sus
            string indicator = "";

            if (indicatorText.ToLower().Contains("toată cartea"))
            {
                // Comentariu pe carte
                var cmdUltimCap = new SQLiteCommand("SELECT MAX(Capitol) FROM Versete WHERE Carte = @carte", conn);
                cmdUltimCap.Parameters.AddWithValue("@carte", carte);
                int ultimCapitol = Convert.ToInt32(cmdUltimCap.ExecuteScalar());

                var cmdUltimVerset = new SQLiteCommand("SELECT MAX(Verset) FROM Versete WHERE Carte = @carte AND Capitol = @capitol", conn);
                cmdUltimVerset.Parameters.AddWithValue("@carte", carte);
                cmdUltimVerset.Parameters.AddWithValue("@capitol", ultimCapitol);
                int ultimVerset = Convert.ToInt32(cmdUltimVerset.ExecuteScalar());

                indicator = $"{carte} 1:1-{ultimCapitol}:{ultimVerset}";
            }
            else if (Regex.IsMatch(indicatorText, @"capitolul\s+\d+\s+versetul\s+\d+", RegexOptions.IgnoreCase))
            {
                // Comentariu pe verset -> afișăm versetul deschis
                indicator = $"{carte} {capitol}:{verset}";
            }
            else if (Regex.IsMatch(indicatorText, @"capitolul\s+\d+", RegexOptions.IgnoreCase))
            {
                // Comentariu pe capitol -> interval complet
                var cmdVersete = new SQLiteCommand("SELECT MAX(Verset) FROM Versete WHERE Carte = @carte AND Capitol = @capitol", conn);
                cmdVersete.Parameters.AddWithValue("@carte", carte);
                cmdVersete.Parameters.AddWithValue("@capitol", capitol);
                int ultimulVerset = Convert.ToInt32(cmdVersete.ExecuteScalar());

                indicator = $"{carte} {capitol}:1-{ultimulVerset}";
            }

            lblComentator.FontFamily = new FontFamily("Segoe UI Emoji");
            lblComentator.Text = $"Comentator: {autor} — {indicator} 📖";
            this.Title = $"Comentariu la {indicator} 📖";

        }










        private void BtnCopiaza_Click(object sender, RoutedEventArgs e)
        {
            var textRange = new TextRange(richComentariu.Document.ContentStart, richComentariu.Document.ContentEnd);
            Clipboard.SetText(textRange.Text.Trim());
        }


        private void BtnInchide_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }


       private void AfiseazaComentariu(string text)
{
    if (flowDoc == null)
    {
        flowDoc = new FlowDocument();
        richComentariu.Document = flowDoc;
    }

    flowDoc.Blocks.Clear();

    // dacă textul e null sau gol, punem mesaj implicit
    var paragrafe = string.IsNullOrWhiteSpace(text)
        ? new[] { "Nu există comentariu pentru acest verset." }
        : text.Split(new[] { "\n" }, StringSplitOptions.None);

    foreach (var paragraf in paragrafe)
    {
        // 🔹 verificăm dacă paragraf conține referințe biblice
        var block = ConvertParagrafCuLinkuri(paragraf.Trim());

        // dacă nu s-a găsit nicio carte biblică, îl punem ca text simplu
        if (block == null)
        {
            var simplu = new Paragraph(new Run(paragraf.Trim()))
            {
                Margin = new Thickness(0, 0, 0, 10)
            };
            flowDoc.Blocks.Add(simplu);
        }
        else
        {
            flowDoc.Blocks.Add(block);
        }
    }
}




        

        private static readonly HashSet<string> CartiBiblice = new HashSet<string>
{
    "GEN","EX","LEV","NUM","DEUT","IOS","JUD","RUT","1SAM","2SAM","1REG","2REG",
    "1CRON","2CRON","EZRA","NEEM","EST","IOB","PS","PROV","ECCL","CANT","ISA","IER",
    "PLNG","EZEC","DAN","OSEA","IOEL","AMOS","OBAD","IONA","MIH","NAUM","HAB","TCEF",
    "HAG","ZAH","MALE","MAT","MAR","LUC","IOAN","FAP","ROM","1COR","2COR","GAL","EFES",
    "FIL","COL","1TES","2TES","1TIM","2TIM","TIT","FILIM","EVR","IAC","1PET","2PET",
    "1IOAN","2IOAN","3IOAN","IUDA","APOC"
};
        private Paragraph ConvertParagrafCuLinkuri(string text)
        {
            var paragraf = new Paragraph { Margin = new Thickness(0, 0, 0, 10) };

            var regex = new Regex(@"((?:[1-3]\s*)?[A-Za-zĂÂÎȘȚăâîșț]+)\.?\s+(\d+:\d+(?:-\d+)?)");


            int lastIndex = 0;
            foreach (Match match in regex.Matches(text))
            {
                if (match.Index > lastIndex)
                {
                    string before = text.Substring(lastIndex, match.Index - lastIndex);
                    paragraf.Inlines.Add(new Run(before));
                }

                string carte = match.Groups[1].Value.Trim();
                string carteNormalizata = NormalizareCarte(carte);

                // 🔹 formatare: prima literă mare, restul mici
                string carteFormata = char.ToUpper(carteNormalizata[0]) + carteNormalizata.Substring(1).ToLower();

                string referinta = $"{carteFormata} {match.Groups[2].Value}";

                var hyperlink = new Hyperlink(new Run(referinta))
                {
                    Foreground = Brushes.Blue,
                    TextDecorations = TextDecorations.Underline,
                    Cursor = Cursors.Hand,
                    ToolTip = $"Deschide comentariul pentru {referinta}"
                };

                hyperlink.Click += (s, e) =>
                {
                    var parts = match.Groups[2].Value.Split(':');
                    int capitol = int.Parse(parts[0]);
                    string versetPart = parts[1];

                    if (versetPart.Contains("-"))
                    {
                        var range = versetPart.Split('-');
                        AfiseazaIntervalVersete(carteNormalizata, capitol,
                            int.Parse(range[0]), int.Parse(range[1]));
                    }
                    else
                    {
                        AfiseazaVersetDinBaza(carteNormalizata, capitol, int.Parse(versetPart));
                    }
                };

                paragraf.Inlines.Add(new Run(" "));
                paragraf.Inlines.Add(hyperlink);
                paragraf.Inlines.Add(new Run(" "));

                lastIndex = match.Index + match.Length;
            }

            if (lastIndex < text.Length)
            {
                string rest = text.Substring(lastIndex);
                paragraf.Inlines.Add(new Run(rest));
            }

            return paragraf;
        }






        private void AfiseazaVersetDinBaza(string cartePrescurtata, int capitol, int verset)
        {
            

            string carte = NormalizareCarte(cartePrescurtata);
           

            var cmd = new SQLiteCommand(
                "SELECT Text, Trimiteri FROM Versete WHERE Carte = @carte AND Capitol = @capitol AND Verset = @verset",
                conn
            );
            

            cmd.Parameters.AddWithValue("@carte", carte);
            cmd.Parameters.AddWithValue("@capitol", capitol);
            cmd.Parameters.AddWithValue("@verset", verset);

            using (var reader = cmd.ExecuteReader())
            {
                if (reader.Read())
                {
                    string textVerset = reader["Text"].ToString();
                    string trimiteri = reader["Trimiteri"].ToString();



                    var fereastra = new FereastraVerset(carte, capitol, verset, textVerset);
                    fereastra.Owner = this;
                    fereastra.ShowDialog();


                    if (!string.IsNullOrWhiteSpace(trimiteri))
                    {
                        
                    }
                }
                else
                {
                    MessageBox.Show($"❌ Nu există versetul {carte} {capitol}:{verset} în baza de date!");
                }
            }
        }
        private void AfiseazaIntervalVersete(string cartePrescurtata, int capitol, int versetStart, int versetEnd)
        {
            string carte = NormalizareCarte(cartePrescurtata);
            var cmd = new SQLiteCommand(
                "SELECT Verset, Text FROM Versete WHERE Carte = @carte AND Capitol = @capitol AND Verset BETWEEN @start AND @end ORDER BY Verset",
                conn
            );
            cmd.Parameters.AddWithValue("@carte", carte);
            cmd.Parameters.AddWithValue("@capitol", capitol);
            cmd.Parameters.AddWithValue("@start", versetStart);
            cmd.Parameters.AddWithValue("@end", versetEnd);

            var listaVersete = new List<(int Verset, string Text)>();

            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    int nrVerset = Convert.ToInt32(reader["Verset"]);
                    string text = reader["Text"].ToString();
                    listaVersete.Add((nrVerset, text));
                }
            }

            var fereastra = new FereastraVerset(carte, capitol, listaVersete);
            fereastra.Owner = this;
            fereastra.ShowDialog();
        }


        private List<Inline> ParseTextCuRosu(string text)
        {
            var inlines = new List<Inline>();
            var regex = new Regex(@"<ROSU>(.*?)<\/ROSU>", RegexOptions.IgnoreCase);

            int lastIndex = 0;
            foreach (Match match in regex.Matches(text))
            {
                if (match.Index > lastIndex)
                {
                    string before = text.Substring(lastIndex, match.Index - lastIndex);
                    inlines.Add(new Run(before));
                }

                string rosuText = match.Groups[1].Value;
                var runRosu = new Run(rosuText)
                {
                    Foreground = Brushes.Red,
                    FontWeight = FontWeights.Bold
                };
                inlines.Add(runRosu);

                lastIndex = match.Index + match.Length;
            }

            if (lastIndex < text.Length)
            {
                string rest = text.Substring(lastIndex);
                inlines.Add(new Run(rest));
            }

            return inlines;
        }







        private string NormalizareCarte(string prescurtare)
        {
            prescurtare = prescurtare
                .ToUpper()
                .Replace(" ", "")
                .Replace(".", "")
                .Replace("-", "")
                .Replace("Ă", "A")
                .Replace("Â", "A")
                .Replace("Î", "I")
                .Replace("Ș", "S")
                .Replace("Ț", "T");

            switch (prescurtare)
            {
                // Vechiul Testament
                case "GEN": return "GENEZA";
                case "EX": return "EXODUL";
                case "LEV": return "LEVITICUL";
                case "NUM": return "NUMERI";
                case "DEUT": return "DEUTERONOMUL";
                case "IOS": return "IOSUA";
                case "JUD": return "JUDECATORI";
                case "RUT": return "RUT";
                case "1SAM": return "1 SAMUEL";
                case "1SAMUEL": return "1 SAMUEL";
                case "2SAM": return "2 SAMUEL";
                case "2SAMUEL": return "2 SAMUEL";
                case "1REG": return "1 IMPARATI";
                case "1IMPARATI": return "1 IMPARATI";
                case "2REG": return "2 IMPARATI";
                case "2IMPARATI": return "2 IMPARATI";
                case "1CRON": case "1CRONICI": return "1 CRONICI";
                case "2CRON": case "2CRONICI": return "2 CRONICI";
                case "EZRA": return "EZRA";
                case "NEEM": return "NEEMIA";
                case "EST": return "ESTERA";
                case "IOB": return "IOB";
                case "PS": case "PSALM": case "PSALMI": return "PSALMII";
                case "PROV": return "PROVERBE";
                case "ECCL": case "ECLEZ": return "ECLESIASTUL";
                case "CANT": return "CANTAREA CANTARILOR";
                case "ISA": return "ISAIA";
                case "IER": return "IEREMIA";
                case "PLNG": case "PLANGERILE": case "PLANGERILELUIEREMIA": return "PLANGERILE LUI EREMIA";
                case "EZEC": return "EZECHEL";
                case "DAN": return "DANIEL";
                case "OSEA": return "OSEA";
                case "IOEL": return "IOEL";
                case "AMOS": return "AMOS";
                case "OBAD": return "OBADIA";
                case "IONA": return "IONA";
                case "MIH": case "MICA": return "MICA";
                case "NAUM": return "NAUM";
                case "HAB": return "HABACUC";
                case "TCEF": case "TEF": return "TEFANIA";
                case "HAG": return "HAGAI";
                case "ZAH": return "ZAHARIA";
                case "MAL": case "MALAHIA": return "MALEAHI";

                // Noul Testament
                case "MAT": return "MATEI";
                case "MAR": return "MARCU";
                case "LUC": return "LUCA";
                case "IOAN": return "IOAN";
                case "FAP": return "FAPTELE APOSTOLILOR";
                case "ROM": return "ROMANI";
                case "1COR": return "1 CORINTENI";
                case "1CORINTENI": return "1 CORINTENI";
                case "2COR": return "2 CORINTENI";
                case "2CORINTENI": return "2 CORINTENI";
                case "GAL": return "GALATENI";
                case "EFES": return "EFESENI";
                case "FIL": return "FILIPENI";
                case "COL": return "COLOSENI";
                case "1TES": return "1 TESALONICENI";
                case "1TESALONICENI": return "1 TESALONICENI";
                case "2TES": return "2 TESALONICENI";
                case "2TESALONICENI": return "2 TESALONICENI";
                case "1TIM": return "1 TIMOTEI";
                case "1TIMOTEI": return "1 TIMOTEI";
                case "2TIM": return "2 TIMOTEI";
                case "2TIMOTEI": return "2 TIMOTEI";
                case "TIT": return "TIT";
                case "FILIM": return "FILIMON";
                case "EVR": return "EVREI";
                case "IAC": return "IACOV";
                case "1PET": return "1 PETRU";
                case "1PETRU": return "1 PETRU";
                case "2PET": return "2 PETRU";
                case "2PETRU": return "2 PETRU";
                case "1IOAN": return "1 IOAN";
                case "1IOANUL": return "1 IOAN";
                case "2IOAN": return "2 IOAN";
                case "2IOANUL": return "2 IOAN";
                case "3IOAN": return "3 IOAN";
                case "3IOANUL": return "3 IOAN";
                case "IUDA": return "IUDA";
                case "APOC": return "APOCALIPSA";

                default: return prescurtare;
            }
        }






    }
}

