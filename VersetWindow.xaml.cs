using System.ComponentModel;
using System.Data.SQLite;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace Biblia
{
    public partial class VersetWindow : Window
    {
        public VersetWindow()
        {
            InitializeComponent();
        }

        public void NavigheazaLaVerset(string carte, int capitol, int verset)
        {
            // Normalizează numele cărții
            carte = NormalizeCarte(carte);


            using var conn = new SQLiteConnection("Data Source=biblia.db");
            conn.Open();

            // Afișează textul versetului
            var cmdVerset = new SQLiteCommand("SELECT Text FROM Versete WHERE LOWER(Carte) = LOWER(@carte) AND Capitol = @capitol AND Verset = @verset", conn);
            cmdVerset.Parameters.AddWithValue("@carte", carte);
            cmdVerset.Parameters.AddWithValue("@capitol", capitol);
            cmdVerset.Parameters.AddWithValue("@verset", verset);

            var textVerset = cmdVerset.ExecuteScalar()?.ToString() ?? "[verset inexistent]";
            tbVerset.Text += $"{carte} {capitol}:{verset}\n{textVerset}\n\n";



            // Afișează explicația (dacă există)
            var cmdExplicatie = new SQLiteCommand("SELECT TextExplicativ FROM ExplicatiiVerset WHERE Carte = @carte AND Capitol = @capitol AND Verset = @verset", conn);
            cmdExplicatie.Parameters.AddWithValue("@carte", carte);
            cmdExplicatie.Parameters.AddWithValue("@capitol", capitol);
            cmdExplicatie.Parameters.AddWithValue("@verset", verset);

            
            
        }

        public void NavigheazaLaTrimitere(string carte, int capitol, int verset, string textTrimitere)
        {
            carte = NormalizeCarte(carte);

            tbVerset.Inlines.Clear();

            tbVerset.Inlines.Add(new Run($"{carte} {capitol}:{verset}\n")
            {
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.DarkSlateGray
            });

            tbVerset.Inlines.Add(new Run(textTrimitere)
            {
                Foreground = Brushes.Black
            });
        }



        public static string NormalizeCarte(string carte)
        {
            if (string.IsNullOrWhiteSpace(carte))
                return carte;

            carte = carte.Trim().ToLower();

            var mapare = new Dictionary<string, string>
            {
                ["gen"] = "GENEZA",
                ["gen."] = "GENEZA",
                ["exod"] = "EXOD",
                ["exod."] = "EXOD",
                ["lev"] = "LEVITIC",
                ["lev."] = "LEVITIC",
                ["num"] = "NUMERI",
                ["num."] = "NUMERI",
                ["deut"] = "DEUTERONOM",
                ["deut."] = "DEUTERONOM",
                ["ios"] = "IOSUA",
                ["ios."] = "IOSUA",
                ["jud"] = "JUDECĂTORI",
                ["jud."] = "JUDECĂTORI",
                ["rut"] = "RUT",
                ["rut."] = "RUT",
                ["1 împ"] = "1 ÎMPĂRAȚI",
                ["1 împ."] = "1 ÎMPĂRAȚI",
                ["2 împ"] = "2 ÎMPĂRAȚI",
                ["2 împ."] = "2 ÎMPĂRAȚI",
                ["1 cron"] = "1 CRONICI",
                ["1 cron."] = "1 CRONICI",
                ["2 cron"] = "2 CRONICI",
                ["2 cron."] = "2 CRONICI",
                ["ezra"] = "EZRA",
                ["ezra."] = "EZRA",
                ["neem"] = "NEEMIA",
                ["neem."] = "NEEMIA",
                ["est"] = "ESTERA",
                ["est."] = "ESTERA",
                ["iov"] = "IOV",
                ["iov."] = "IOV",
                ["ps"] = "PSALMII",
                ["ps."] = "PSALMII",
                ["prov"] = "PROVERBE",
                ["prov."] = "PROVERBE",
                ["ecle"] = "ECLESIASTUL",
                ["ecle."] = "ECLESIASTUL",
                ["cant"] = "CÂNTAREA CÂNTĂRILOR",
                ["cant."] = "CÂNTAREA CÂNTĂRILOR",
                ["isa"] = "ISAIA",
                ["isa."] = "ISAIA",
                ["ier"] = "IEREMIA",
                ["ier."] = "IEREMIA",
                ["plâng"] = "PLÂNGERILE",
                ["plâng."] = "PLÂNGERILE",
                ["ezec"] = "EZECHIEL",
                ["ezec."] = "EZECHIEL",
                ["dan"] = "DANIEL",
                ["dan."] = "DANIEL",
                ["os"] = "OSEA",
                ["os."] = "OSEA",
                ["ioel"] = "IOEL",
                ["ioel."] = "IOEL",
                ["amos"] = "AMOS",
                ["amos."] = "AMOS",
                ["obad"] = "OBADIA",
                ["obad."] = "OBADIA",
                ["iona"] = "IONA",
                ["iona."] = "IONA",
                ["mica"] = "MICA",
                ["mica."] = "MICA",
                ["nah"] = "NAUM",
                ["nah."] = "NAUM",
                ["hab"] = "HABACUC",
                ["hab."] = "HABACUC",
                ["țef"] = "ȚEFANIA",
                ["țef."] = "ȚEFANIA",
                ["hag"] = "HAGAI",
                ["hag."] = "HAGAI",
                ["zac"] = "ZAHARIA",
                ["zac."] = "ZAHARIA",
                ["mal"] = "MALEAHI",
                ["mal."] = "MALEAHI",
                ["Matei "] = "MATEI",
                ["mat."] = "MATEI",
                ["marc"] = "MARCU",
                ["marc."] = "MARCU",
                ["Luca "] = "LUCA",
                ["luc."] = "LUCA",
                ["ioan"] = "IOAN",
                ["ioan."] = "IOAN",
                ["fapte"] = "FAPTELE APOSTOLILOR",
                ["fapte."] = "FAPTELE APOSTOLILOR",
                ["rom"] = "ROMANI",
                ["rom."] = "ROMANI",
                ["1 cor"] = "1 CORINTENI",
                ["1 cor."] = "1 CORINTENI",
                ["2 cor"] = "2 CORINTENI",
                ["2 cor."] = "2 CORINTENI",
                ["gal"] = "GALATENI",
                ["gal."] = "GALATENI",
                ["efes"] = "EFESENI",
                ["efes."] = "EFESENI",
                ["filip"] = "FILIPENI",
                ["filip."] = "FILIPENI",
                ["col"] = "COLOSENI",
                ["col."] = "COLOSENI",
                ["1 tes"] = "1 TESALONICENI",
                ["1 tes."] = "1 TESALONICENI",
                ["2 tes"] = "2 TESALONICENI",
                ["2 tes."] = "2 TESALONICENI",
                ["1 tim"] = "1 TIMOTEI",
                ["1 tim."] = "1 TIMOTEI",
                ["2 tim"] = "2 TIMOTEI",
                ["2 tim."] = "2 TIMOTEI",
                ["tit"] = "TIT",
                ["tit."] = "TIT",
                ["fil"] = "FILIMON",
                ["fil."] = "FILIMON",
                ["evr"] = "EVREI",
                ["evr."] = "EVREI",
                ["iac"] = "IACOV",
                ["iac."] = "IACOV",
                ["1 pet"] = "1 PETRU",
                ["1 pet."] = "1 PETRU",
                ["2 pet"] = "2 PETRU",
                ["2 pet."] = "2 PETRU",
                ["1 ioan"] = "1 IOAN",
                ["1 ioan."] = "1 IOAN",
                ["2 ioan"] = "2 IOAN",
                ["2 ioan."] = "2 IOAN",
                ["3 ioan"] = "3 IOAN",
                ["3 ioan."] = "3 IOAN",
                ["iuda"] = "IUDA",
                ["iuda."] = "IUDA",
                ["apoc"] = "APOCALIPSA",
                ["apoc."] = "APOCALIPSA"
            };



            // Normalizează forma
            if (mapare.ContainsKey(carte))
                return mapare[carte];

            // Dacă nu e mapată, capitalizează prima literă
            return char.ToUpper(carte[0]) + carte.Substring(1);
        }


    }
}

