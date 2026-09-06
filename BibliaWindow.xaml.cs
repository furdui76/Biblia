using Biblia;
using Biblia.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Data.Sqlite;
using Microsoft.Win32;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Globalization;
using System.IO;
using System.Linq; // pentru FirstOrDefault
using System.Net;
using System.Net.Mail;
using System.Speech.Synthesis;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Navigation;
using System.Windows.Threading;
using OxmlBold = DocumentFormat.OpenXml.Wordprocessing.Bold;
using OxmlBreak = DocumentFormat.OpenXml.Wordprocessing.Break;
using OxmlColor = DocumentFormat.OpenXml.Wordprocessing.Color;
using OxmlFontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;
using OxmlIndentation = DocumentFormat.OpenXml.Wordprocessing.Indentation;
using OxmlJustification = DocumentFormat.OpenXml.Wordprocessing.Justification;
using OxmlJustificationValues = DocumentFormat.OpenXml.Wordprocessing.JustificationValues;
using OxmlPageMargin = DocumentFormat.OpenXml.Wordprocessing.PageMargin;
using OxmlPageSize = DocumentFormat.OpenXml.Wordprocessing.PageSize;
// Aliasuri pentru OpenXML
using OxmlParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using OxmlParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using OxmlRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using OxmlRunFonts = DocumentFormat.OpenXml.Wordprocessing.RunFonts;
using OxmlRunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using OxmlSectionProperties = DocumentFormat.OpenXml.Wordprocessing.SectionProperties;
using OxmlSpacingBetweenLines = DocumentFormat.OpenXml.Wordprocessing.SpacingBetweenLines;
using OxmlText = DocumentFormat.OpenXml.Wordprocessing.Text;
using SWD = System.Windows.Documents;
using WordRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using WpfBorder = System.Windows.Controls.Border;
using WpfFontFamily = System.Windows.Media.FontFamily;
using WpfHyperlink = System.Windows.Documents.Hyperlink;
// Aliasuri WPF
using WpfRun = System.Windows.Documents.Run;
using WpfTextAlignment = System.Windows.TextAlignment;
using Run = System.Windows.Documents.Run;
using System.Text.RegularExpressions;












namespace Biblia
{


    public class SchitaPredica
    {
        public string Carte { get; set; } = string.Empty;
        public int Capitol { get; set; }
        public string Titlu { get; set; } = string.Empty;
        public List<VersetSchita> Versete { get; set; } = new();
    }






    public partial class BibliaWindow : Window
    {
        private SpeechSynthesizer sintetizatorGlobal;
        private string? textCititComplet;
        private int pozitieCurenta;
        private bool esteInPauza = false;
        private bool modSchitaActiv = false;
        public bool EsteModSchitaActiv => modSchitaActiv;
        public string CarteCurenta => versetCarteCurenta;
        public int CapitolCurent => versetCapitolCurent;
        private List<ListBoxItem> listaNormalaVersete = new List<ListBoxItem>();
        private bool esteInModCautare = false;
        private bool modComentariiActiv = false;
        private string fontSchita = "Georgia";
        private double marimeFontSchita = 12;
        private SQLiteConnection conn;
        private bool trimiteriVizibile = false;
        private string versetCarteCurenta;
        private Popup popup;
        private int versetCapitolCurent;
        private Brush culoareTextVerset => modIntunecatActiv ? Brushes.White : Brushes.Black;
        private List<VersetSchita> schitaCurenta = new();
        private bool isDraggingSchita = false;
        private Point startPointSchita;
        private bool istoricActiv = false;
        // câmp global în clasa ta
        private bool esteSchitaActivata = false;
        private string cuvantCurentCautat = string.Empty;
        // statistici pentru căutare/încărcare
        private int totalAparitii = 0;
        private int totalVerseteGasite = 0;
        private SQLiteConnection conexiune;
        private Regex regex;
      private static Window? ultimaFereastraVerset;
        private ModAfisare modCurent = ModAfisare.Scriptura;
        private void SeteazaMod(ModAfisare mod)
        {
            modCurent = mod;
        }
        private bool EsteMod(ModAfisare mod)
        {
            return modCurent == mod;
        }
        private bool arataStele = false;


        private Dictionary<string, List<VersetSchita>> favoritePeListe = new();

        public BibliaWindow()
        {
            InitializeComponent();

            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "biblia.db");
            conn = new SQLiteConnection($"Data Source={path}");
            conn.Open();

            // 🔥 Întâi generăm butoanele
            LoadCartiCuButoane();

         

            // Continuăm cu inițializările
            favoritePeListe = IncarcaFavoriteDinBazaDeDate();
            comboListeFavorite.ItemsSource = favoritePeListe.Keys.ToList();
            SeteazaMod(ModAfisare.Scriptura);
            versetCarteCurenta = null;
            versetCapitolCurent = 0;
            runTitlu.Text = "📖 Cărțile BIBLIEI";
            runStea.Text = "";
        }















        // Dicționar cu descrieri istorice pentru fiecare carte
        Dictionary<string, string> istoricCarti = new Dictionary<string, string>
{
    // Vechiul Testament

    { "GENEZA", " 🏛️ <titlu>Istoric</titlu>\r\n✍️ <autor>Autoru cărți Moise</autor>\r\n📜 <explicație> Explicația Cărții pentru Geneza</explicație>\r\n\r\n Titlul din limba romana provine de la cuvântul latin „genesis\" care înseamnă „început\", în limba ebraica, cartea aceasta poarta numirea primelor cuvinte din debutul ei: „La început („beresit). Acest titlu este cum nu se poate mai nimerit pentru o carte în care Dumnezeu a binevoit sa ne dea informații cu privire la începuturile tuturor lucrurilor din lumea necunoscuta în care deschidem ochii prin nașterea biologica.\r\n\r\n Autorul: Cartea aceasta a fost scrisa de Moise si, împreuna cu următoarele patru scrise de el,"+" face parte dintr-o grupa de cinci cărți cunoscute sub numirea de: „Pentateuhul lui Moise.Dovezi în sprijinul părerii ca Moise a fost autorul acestor cărți se găsesc în: (1) textul acestor cărți Exod 17:14; Exod 24:4-7 Exod 34:27; Num. 33:1-2; Deut. 31:9, (2) alte cărți ale Vechiului Testament (Iosua 1:7-8; Iosua 8:32, (Iosua 22:5,) ( 1Împ. 2:3) (2Împ. 14:6) Ezra 6:18; Dan. 9:11-13; Mal. 4:4), (3) în cărțile Noului Testament (Mat. 19:8; Marc 12:26; Ioan 5:46-47;Ioan 7:19; Rom. 10:5), (4) faptul ca.descrierile apar a fi ale unui martor ocular, nu a unuia care ar fi trăit la multe secole după aceea (Exod 15:27; Num 21:1-31;Num 11:7-8), (5) faptul ca autorul manifesta o cunoaștere foarte exacta asupra numelor, cuvintelor, obiceiurilor si locurilor din Egipt (Gen 13:10; Gen 16:1-3; Gen 33:18; Gen 41:43 Fapte 7:22)\r\n\r\n Conținutul: Cartea Genezei conține informații despre „începuturile multor lucruri din lumea în care trăim noi astăzi. Descrierea acestor „începuturi este făcută concentric începând cu informații globale despre realitatea creației si apropiindu-se apoi treptat, treptat de lumea imediat înconjurătoare omului. Aceasta apropiere cosmogonica sugerează ca întreaga carte este un răspuns pe care Dumnezeu îl da celor mai arzătoare întrebări ale ființei umane: Cine sânt eu? Ce este lumea aceasta în care m-am născut? Cine a făcut-o? De ce m-am născut si care este rațiunea existentei mele?\r\n  Textul Genezei apare de 60 de ori în conținutul Noului Testament dovedind într-un sens profund ca germenii tuturor revelațiilor Dumnezeiești sânt plantate aici si ca cine vrea sa înțeleagă pe deplin aceasta revelație trebuie sa înceapă cu aceasta carte.\r\n Geneza este o carte istorica. Ea cuprinde evenimente din „protoistoria omenirii\". De fapt am putea împărții întreaga Biblie în doua secțiuni istorice distincte. Evenimentele dinaintea căderii omului (Gen 1:1 până la Gen 3:24) si evenimentele de după aceea (Gen 4:1 până la sfârșitul Bibliei). În cea de a doua secțiune istorica sânt redate evenimentele din planul de „mântuire\" a ființei umane, adică acele evenimente cauzate de Dumnezeu sa se întâmple pentru salvarea omului de la pierzarea veșnica.\r\n Sânt multe feluri în care se poate face istorie; exista o istorie a mijloacelor de producție (tehnologie), o alta a formelor diferite de alcătuire a societății umane (sociologie), o alta a diferitelor războaie sau a cuceririlor teritoriale, o alta istorie a cuceririi spațiului înconjurător si încă o alta a abilitații omului de a-si explica lumea si rațiunea ei de existenta (filosofie). Biblia este o carte de istorie în măsura în care cuprinde evenimentele cauzate de Dumnezeu pentru salvarea oamenilor de la pierzarea veșnica produsa de neascultarea din gradina Edenului, în lumina destinului nostru veșnic, Biblia, si implicit cartea Genezei, este singura narațiune istorica care ne poate influenta soarta noastră eterna. Ignorarea celorlalte cărți de istorie ne poate dauna relativ, dar ignorarea Bibliei ne poate atrage suferința noastră absoluta si de neînlăturat.\r\n În mod normal, după căderea omului în neascultare, Dumnezeu a trebuit sa exercite pedeapsa si sa refacă echilibrul dreptății. Dumnezeu a ales însă sa nu ne facă ceva rău care sa ne distrugă veșnic; pedepsele Lui sânt limitate în măsura si educative în scop. Hotărârea lui Dumnezeu de a restrânge „dimensiunile\" pedepsei pentru a ne face bine este ilustrata în cele 8 „legăminte\" pe care El le-a făcut cu omul în decursul istoriei. Primele 4 legăminte din aceasta serie de 8 sânt redate în cuprinsul cârtii Geneza. Iată însă o lista completa a tuturor celor 8:\r\n (1) Edenic  (Gen. 2:16),\r\n (2) Adamic  (Gen. 3:15),\r\n (3) prin Noe  (Gen. 9:16),\r\n (4) Avraamic  (Gen. 12:2),\r\n (5) Mozaic  (Gen. 19:5),\r\n (6) Palestinian  (Deut. 30:3),\r\n (7) Davidic  (2Sam. 7:16)\r\n si (8) Legământul cel Nou   (Evr 8:8; (1 Cor. 11:23-25;)\r\n Așezata la începutul celor 66 de cărți ale Bibliei, Geneza este si începutul „revelației de Sine\" a lui Dumnezeu, a descoperirii ființei Lui către făptura creată. Aceasta descoperire treptata este progresiva în esența si culminează în Cristos, Dumnezeu adevărat făcut trup si coborât printre oameni (Ioan 1:18; Ioan 14:8-10;\" Col. 1:15;\r\n  Col. 2:9;).\r\n Cunoașterea pe care oamenii au primit-o despre Dumnezeu este exprimata în „numele\" cu care L-au identificat. În mod absolut, Dumnezeu nu poate avea un nume deoarece orice cuvânt reprezintă o realitate statica, iar Dumnezeu este infinit în dimensiunile ființei Sale si în posibilitățile Lui de manifestare. Totuși, cartea Geneza ne pune la dispoziție primele trei numiri cu care L-au identificat oamenii pe Dumnezeu: Elohim, Iehova si Adonai, date oamenilor într-o ordine prestabilita si desăvârșita care nu poate fi schimbata fără a altera caracterul revelației lui Dumnezeu către oameni (Exod 6:3), precum si cele cinci numiri compuse ale numelui lui Dumnezeu: El-Elion: Dumnezeul Cel Prea înalt (Gen. 14:18), Iehova Elohim: Domnul Dumnezeu (Gen. 15:2), El-Roi: Dumnezeul care mă vede (Gen. 16:13), Iehova-Iire: Domnul va purta de grija (Gen. 22:14) El-Elohe-Israel: Domnul este Dumnezeul lui Israel (Gen. 33:20).\r\n Elohim este o forma de plural care denumește un aspect caracteristic al ființei divine: Dumnezeu întreit în ființa (Gen. 1:26), dar unic în natura si persoana (Deut. 6:4)\r\n Iehova nu este un nume propriu zis, ci un grup de cuvinte care desfășoară ceva din caracterul infinit al prezentei si posibilităților divine: „Eu sânt Cel ce sânt\" (Exod 3:11-15). Ocazia în care a apărut acest nume i-a făcut pe mulți sa se oprească asupra lui ca asupra singurului nume veritabil al dumnezeirii. Este bine de știut însă ca evreii au pierdut ei înșiși pronunțarea corecta a acestei perifraze notate în textul vechi doar prin consoanele HH si semivocalele Y, W (lingviștii de astăzi înclina totuși spre forma Yahweh). În timpul Evului Mediu, rabinii au înlocuit vocalele originale cu vocalele din cel de al treilea nume pentru divinitate: Adonai si astfel s-a ajuns la numirea de „Iehova\". Astăzi nimeni nu poate spune ca știe cum a sunat vocea din ceruri prin care S-a identificat Dumnezeu înaintea lui Moise, iar faptul ca în textul Bibliei au apărut apoi alte numiri pentru Dumnezeu ne arata clar ca nu este bine sa ne legam prea mult de aceasta singura numire.\r\n Adonai, sau mai scurt „Adon\" este numirea veche pentru „Stăpân\" sau „Domn\" si este aplicata în textul Vechiului Testament si omului si divinității (Gen. 15:2). Atunci când este aplicat oamenilor „Adonai\" se scrie cu litera mica.\r\n Faptul ca Geneza începe cu cuvintele acestea: „La început Dumnezeu a făcut cerurile si pământul\" este de o importanta covârșitoare. Aceasta afirmație limitează rătăcirile imaginației noastre care-l caută pe Dumnezeu în forme aberante si arbitrare.\r\n„ La început Dumnezeu...” - neaga totodată si Ateismul si Politeismul. Între vidul nimicului si confuzia mulțimii de dumnezei, Geneza ni-L așază înainte pe maiestuosul Dumnezeu unic care este sursa, suportul si scopul tuturor lucrurilor.\r\n„ La început Dumnezeu a făcut...” - neaga Fatalismul cu învățătura lui bazata pe hazard.\r\n„ La început Dumnezeu a făcut...” - neaga si Evoluționismul haotic, mașinist si impersonal.\r\n„ La început Dumnezeu a făcut Cerurile si pământul - neaga Panteismul care susține ca Dumnezeu este totuna cu creația Sa, dar neaga si Materialismul care, amețind-ne cu ceea ce se vede si pretinzând ca materia este veșnica, caută sa ne ascundă izvorul creator, începutul si Judecătorul desăvârșit al tuturor celor create.\r\n Împărțirea Cârti \r\n Tema Genezei fiind: „Suveranitatea divina în creație, istorie si mântuire\", am putea sistematiza cuprinsul cârtii după cum urmează: \r\n I. Protoistoria Omenirii \r\n Patru evenimente de seama\r\n Creația - Suveranitatea divina asupra lumii fizice. Prioritățile veșnice ale lui Dumnezeu.\r\n Căderea - Suveranitatea divina în problemele umane. Autoritatea morala divina.\r\n POTOPUL - Suveranitatea divina în pedepsire. Autoritatea judiciara suprema.\r\n CRIZA DE LA BABEL - Suveranitatea divina în răspândirea si specificul raselor. Autoritate sociala suprema. Stăpânirea deplina asupra societății.\r\n II. Istoria patriarhilor\r\n Patru persoane de seama\r\nAVRAAM - Suveranitatea divina în alegere. Chemare supranaturala.\r\n ISAAC - Suveranitatea divina în alegere. Nașterea supranaturala.\r\n IACOV - Suveranitatea divina în alegere. Îngrijire supranaturala.\r\n IOSIF - Suveranitatea divina în călăuzire. Călăuzire supranaturala.\r\n\r\n   📝 Schițai Cârti \r\n\r\n I. FACEREA LUMII Gen 1:1-31, Gen  2:1-25, \r\na. Facerea universului, Gen 1:1, \r\n b. Facerea pământului, Gen 1:2-31, Gen  2:3, \r\n c. Facerea Omului, Gen 2:4-25,\r\n II. Căderea OMULUI Gen 3:1-24,\r\na. Ispitirea, Gen 3:1-7, \r\n b. Pedepsele, Gen 3:8-24, \r\n III. Cain si urmașii lui, Gen 4:1-24, \r\n b. Set, Gen 4:25-26,\r\n c. De la Adam la Noe, Gen 5:1-32, \r\n IV. Viată LUI NOE de la Cap 6 la Cap 10 \r\n a. Cauzele potopului, Gen 6:1-13, \r\n b. Potopul, Gen 6:14-22, \r\nIntrarea in Corabie Gen 7:1-24,\r\nSfârșitul Potopului Gen 8:1-18,\r\n  Evenimente de după potop, Gen 8:18-22, \r\n d. Urmașii lui Noe, Gen 10:1-32,\r\n TURNUL BABEL, Gen 11:1-9, \r\n VI. Viată LUI AVRAAM Gen11:10-25, Strămoșii lui \r\n\r\n📋 Detaliat \r\n\r\n1 Șase Zile ale Creației și Sabatul (Gen 1:1-31 Gen 2:1-4)\r\n2 Grădina Edenului (Gen 2:4-17)\r\n3 Creația Femeii (Gen 2:18-25)\r\n4 Primul păcat și pedeapsa lui (Gen 3:1-24)\r\n5 Cain îl ucide pe Abel (Gen 4:1-26)\r\n6 Descendenții lui Adam la Noe și fiii lui (Gen 5:1-32)\r\n7 Marele Potop (Gen 6:1-8)\r\n8 Legământul cu Noe (Gen 9:1-17)\r\n9 Noe și Fiii Săi (Gen 9:18-29)\r\n10 Națiuni Descinzătoare din Noe (Gen 10:1-32)\r\n11 Turnul Babel (Gen 11:1-9)\r\n12 Descendenții lui Sem (Gen 11:10-26)\r\n13 Descendenții Terei (Gen 11:27-32)\r\n14 Chemarea lui Abram (Gen 12:1-9)\r\n15 Abram și Sarai în Egipt (Gen 12:10-20)\r\n16 Abram și Lot se separă (Gen 13:1-18)\r\n17 Captivitatea lui Lot (Gen 14:1-12)\r\n18 Abram și Melhisedec (Gen 14:13-24)\r\n19 Legământul lui Dumnezeu cu Abram (Gen 15:1-21)\r\n20 Nașterea lui Ismael (Gen 16:1-16)\r\n21 Semnul legământului (Gen 17:1-27)\r\n22 Un fiu promis lui Avraam și Sara (Gen 18:1-15)\r\n23 Judecata rostită asupra Sodomei (Gen 18:16-33)\r\n24 Depravarea Sodomei (Gen 19:1-11)\r\n25 Sodoma și Gomorra distruse (Gen 19:12-29)\r\n26 Originea rușinoasă a Moab și Ammon (Gen 19:30-38)\r\n27 Avraam și Sara la Gerar (Gen 20:1-18)\r\n28 Nașterea lui Isaac (Gen 21:1-8)\r\n29 Hagar și Ismael trimiși departe (Gen 21:9-21)\r\n30 Legământul lui Abram și Abimelec (Gen 21:22-34)\r\n31 Porunca de a jertfi pe Isaac (Gen 22:1-19)\r\n32 Copiii lui Nahor (Gen 22:20-24)\r\n33 Moartea și înmormântarea Sarei (Gen 23:1-20)\r\n34 Nunta lui Isaac și Rebeca (Gen 24:1-67)\r\n35 Avraam se căsătorește cu Ketura (Gen 25:1-6)\r\n36 Moartea lui Avraam (Gen 25:7-11)\r\n37 Descendenții lui Ismael (Gen 25:12-18)\r\n38 Nașterea și tinerețea lui Esau și Iacov (Gen 25:19-28)\r\n39 Esau își vinde dreptul de naștere (Gen 25:29-34)\r\n40 Isaac și Abimelec (Gen 26:1-14)\r\n41 Necazuri cu Fântânii, Legământul cu Dumnezeu (Gen 26:15-25)\r\n42 Legământul cu Abimelec (Gen 26:26-33)\r\n43 Soțiile hitite ale lui Esau (Gen 26:34-35)\r\n44 Iacov fură Binecuvântare (Gen 27:1-40)\r\n45 Rebeka recomandă ca Iacov să fugă (Gen 27:41-45)\r\n46 Iacov scapă de furia lui Esau (Gen 27:46-28:9)\r\n47 Visul lui Iacov la Betel (Gen 28:10-22)\r\n48 Iacov o întâlnește pe Rahela (Gen 29:1-14)\r\n49 Iacov se căsătorește cu fiicele lui Laban (Gen 29:14c-30)\r\n50 Copiii lui Iacov (Gen 29:31-30:24)\r\n51 Iacov prosperă pe seama lui Laban (Gen 30:25-43)\r\n52 Iacov fuge cu familia și turmele (Gen 31:1-21)\r\n53 Laban îl învinge pe Iacov (Gen 31:22-42)\r\n54 Laban și Iacov fac un legământ (Gen 31:43-54)\r\n55 Mahanaim (Gen 32:1-13)\r\n56 Iacov trimite daruri pentru a-l împăca pe Esau (Gen 32:14-22)\r\n57 Iacov se luptă la Peniel (Gen 32:23-33)\r\n58 Iacov și Esau se întâlnesc (Gen 33:1-20)\r\n59 Violul Dinai (Gen 34:1-31)\r\n60 Iacov se întoarce la Betel (Gen 35:1-8)\r\n61 Legământul reînnoit (Gen 35:9-15)\r\n62 Decese și nașterea unui copil (Gen 35:16-29)\r\n63 Descendenții lui Esau (Gen 36:1-43)\r\n64 Iosif visează măreție (Gen 37:1-11)\r\n65 Iosif este vândut de frații săi (Gen 37:12-36)\r\n66 Iuda și Tamara (Gen 38:1-30)\r\n67 Iosif și soția lui Potifar (Gen 39:1-23)\r\n68 Iosif interpretează visele (Gen 40:1-41:45)\r\n69 Copiii lui Iosif (Gen 41:46-57)\r\n70 Prima călătorie a fraților lui Iosif în Egipt (Gen 42:1-38)\r\n71 A doua călătorie a fraților lui Iosif în Egipt (Gen 43:1-45:28)\r\n72 Iacov își aduce întreaga familie în Egipt (Gen 46:1-7)\r\n73 Lista familiei lui Iacov (Gen 46:8-27)\r\n74 Iacov și Faraonul (Gen 46:28-47:12)\r\n75 Foametea din Egipt (Gen 47:13-26)\r\n76 Ultimele zile ale lui Iacov (Gen 47:27-31)\r\n77 Iacov binecuvântează fiii lui Iosif (Gen 48:1-22)\r\n78 Ultimele cuvinte ale lui Iacov către fiii săi (Gen 49:1-28)\r\n79 Moartea și înmormântarea lui Iacov (Gen 49:29-33 Gen 50:1-14)\r\n80 Iosif iartă frații Săi (Gen 50:15-21)\r\n81 Ultimele zile și moartea lui Iosif (Gen 50:22-26)" },
    //EXODL
    { "EXODUL", "Explicatia Cartii: Exod\r\n\r\nExodul urmează Genezei în acelaşi fel în care Noul Testament urmează Vechiului Testament. Geneza ne vorbeşte despre căderea omului la toate examenele şi în toate privinţele; Exodul este descrierea glorioasei intervenţii a Dumnezeului pornit în răscumpărarea creaturii pierdute. Ea ne vorbeşte despre un Dumnezeu suveran care a hotărît în bunăvoinţa Lui să mîntuiască.\r\n\r\nExodul este prin excelenţă cartea care ne vorbeşte mai clar decît oricare alta din Vechiul Testament despre lucrarea de mîntuire. Acţiunea ei debutează în plină tragedie umană, dar se termină cu un triumf al slavei; începe prin a ne spune cum Dumnezeu s-a coborît plin de har în scena lumii pentru a izbăvi un popor pe care-l iubeşte şi se termină cu declaraţia că Dumnezeu a coborît în norul slavei pentru a locui în mijlocul acestui popor răscumpărat.\r\n\r\nTitlul: Titlul românesc vine dintr-o traducere a Bibliei evreieşti în limba greacă numită „Septuaginta\". Exodul sau „Ieşirea\" este o numire cum nu se poate mai nimerită pentru această carte în care Dumnezeu ne relatează despre ieşirea poporului ales din ţara în care suferiseră crunt ca robi timp de cîteva generaţii. Termenul „exod\" se află scris ca atare în 19:1, iar în varianta lui grecească apare în cuprinsul Noului Testament în Evrei 11:22, Luca 9:31 şi 2 Petru 1:15 (în ultimele două texte cu sensul de ieşire ca o „trecere dincolo\" în altă realitate).\r\n\r\nAutorul: Încă din timpul lui Iosua, cartea i-a fost atribuită lui Moise (Iosua 8:31-35; cf. Exodul 20:25), părere confirmată apoi şi de Domnul Isus (Marcu 12:26). Datele cărţii ne fac să credem că autorul ei a fost într-adevăr un om cu o foarte aleasă educaţie, care a locuit în Egipt un timp îndelungat şi a fost el însuşi martor ocular la evenimentele descrise. Autorul a fost familiarizat cu ciclul recoltelor din valea inferioară a Nilului (Exodul 9:31-32); descrierile lui sînt foarte precise şi clare (Exodul 2:3, 12), dîndu-ne amănunte pe care numai un martor ocular putea să le observe (Exodul 15:27).\r\n\r\nData scrierii: Probabil între anii 1445 -1440 î.Cr.\r\n\r\nSubiectul: Tema acestei cărţi este eliberarea din robia egipteană, ca o împlinire a promisiunii făcute în Geneza 15:13-14. Această izbăvire a fost îndeplinită în urma unor intervenţii miracuIoase din partea lui Dumnezeu (plăgile din Exodul 7:14 -12:36) şi n-a cerut din partea evreilor decît o încredere totală în eficacitatea sîngelui simbolic vărsat (Exodul 12:1-13). Ca şi în Noul Testament, eliberarea (sau mîntuirea) are ca scop aducerea celor eliberaţi într-o stare de strînsă părtăşie cu Dumnezeu. După realizarea ieşirii din Egipt, poporului evreu îi este dată Legea, împreună cu o mulţime de alte descoperiri despre adevărata închinăciune înscrisă în tiparul şi simbolurile Cortului întîlnirii. La toate acestea sînt adăugate reguli care reglementează apropierea de Dumnezeu prin intermediul jertfelor practicate de preoţi.\r\n\r\nÎn cartea Exodul, Dumnezeu, care se apropiase pînă atunci de popor pe baza legămîntului Avraamic, îi aşează pe evrei aproape de Sine însuşi prin prerogativele unui alt legămînt: cel Mozaic.\r\n\r\nLegămîntul Avraamic dat în Gen. 12:2-3 fusese întreit în semnificaţie: (1) anunţa transformarea lui Avraam într-un popor foarte mare, (2) îi promitea lui Avraam patru binecuvîntări personale şi (3) prin evrei, făgăduia neamurilor trei lucruri distincte: a. „Voi binecuvînta pe cei ce te vor binecuvînta\", b. „voi blestema pe cei ce te vor blestema\" şi c. toate familiile pămîntului vor fi binecuvîntate în tine\" (aceasta din urmă fiind încă o confirmare a „proto-evengheliei\" din Gen.3:15 care vestise deja hotărîrea că Cel care va veni din sămînţa femeii (Isus) va zdrobi capul şarpelui (Satan).\r\n\r\nLegămîntul Mozaic era cuprins tot în trei secţiuni distincte, după cum urmează: (1) Legea lui Dumnezeu despre neprihănire (Exodul 20:1-26), (2) legile pentru judecată, hotărînd în problemele vieţii sociale ale poporului Israel (Exodul 21:1 - 24:11) şi (3) poruncile care reglementau aspectele închinăciunii religioase a poporului Israel (Exodul 24:12 - 31:18). Aceste trei elemente formează împreună ceea ce Noul Testament numeşte mai tîrziu generic „Legea\" (Matei 5:17, 18).\r\n\r\nRelaţia dintre legămîntul Avraamic şi legămîntul Mozaic este analizată în profunzime de autorul epistolei către Galateni. Legea mozaică a fost „aducătoare de moarte\" (2 Cor. 3:7-8), poruncile, în rînduielile lor ofereau poporului în persoana Marelui Preot un mijlocitor care să facă ispăşirea pentru ei, iar în mulţimea de jertfe era ascunsă „acoperirea\" păcatelor pînă la împlinirea vremii, în anticiparea Crucii mîntuitoare a Domnului Isus (Evrei 5:1-3; 9:6-9; Rom. 3:25-26). Astăzi, creştinul nu se mai află sub acest legămînt Mozaic al faptelor, ci a trecut sub prerogativele pline de har ale Legămîntului celui Nou (Rom. 3:21-27; 6:14-15; Gal. 2:16; 3:10-14, 16-18, 24-26; 4:21-31; evrei 10:11-17).\r\n\r\nLegămîntul Mozaic nu a desfinţat legămîntul Avraamic, ci doar i-a adăugat ceva temporar, menit să aibă valabilitate doar pînă la venirea acelei „seminţe\" care va aduce deplina împăcare (Gal. 3:17-19). Prin intermediul legilor şi poruncilor, Dumnezeu le-a arătat oamenilor normele neprihănirii şi dreptăţii divine. Experienţa trăirii sub pretenţiile Legii i-a convins pe evrei de natura lor păcătoasă, în timp ce preoţia şi mulţimea de jertfe trebuia să le vorbească despre o cale de ieşire de sub imperiul vinovăţiei spre o viaţă de iertare, curăţire, restaurare a părtăşie şi adevărată închinare.\r\n\r\nCuvinte cheie şi terne caracteristice: Cartea Exodul ne aşează înainte cîteva „antetipuri\" ale lui Isus Cristos: Moise (2:2), Paştele (12:11), mana (16:35), stînca (17:6) şi Cortul întîlnirii (25:9).\r\n\r\nExodul ne pune înainte o lectură fascinantă. Există oare în toată istoria un spectacol mai măreţ decît această „scoatere\" a poporului evreu din robia Egiptului? Există oare în oricare alta din celelalte religii ale lumii o privelişte mai maiestuoasă şi mai slăvită decît această apariţie a lui Dumnezeu pe muntele Sinai cu ocazia dării Legii? Care clădire dintre capodoperele arhitectonice ale lumii se poate asemăna cu armonia şi bogăţia de semnificaţii de veşnică importanţă care este înscrisă în structura Cortului întîlnirii? Puteţi găsi undeva o figură mai impozantă ca acest gigantic întemeietor de neam cu numele de Moise? Există în toată istoria biblică o altă epocă mai importantă pentru definirea şi formarea „teocraţiei\"?\r\n\r\nToate aceste evenimente şi date unice se găsesc aici în aceste 40 de capitole ale Exodului care ne vorbesc despre originea, Legea şi organizarea religioasă a poporului Israel.\r\n\r\nSCHIŢA CĂRŢII\r\n\r\nI. EXODUL 1:1 - 18:27\r\na. Suferinţele Israeliţilor, 1:1-22\r\nb. Pregătirea lui Moise, 2:1-25\r\nc. Chemarea lui Moise, 3:1 - 4:31\r\nd. Trimiterea lui Moise la Faraon, 5:1-7:13\r\ne. Prima plagă: sîngele, 7:14-25\r\nf. A doua plagă: broaştele, 8:1-15\r\ng. A treia plagă: păduchii, 8:16-19\r\nh. A patra plagă: muştele, 8:20-32\r\ni. A cincea plaga: ciuma vilelor, 9:1-7\r\nî. A şasea plagă: vărsatul negru, 9:8-12\r\nj. A şaptea plagă: grindina, 9:13-35\r\nk. A opta plagă: lăcustele, 10:1-20\r\nl. A noua plagă: întunerecul, 10:21-29\r\nm. A zecea plagă: moartea, 11:1 -12:36\r\nn. Plecarea din Egipt, 12:37-51\r\no. Paştele şi întîii născuţi, 13:1-16\r\np. Călătoria prin pustie, 13:17-22\r\nr. Trecerea Mării Roşii, 14:1 -15:21\r\ns. Apele de la Mara, 15:22-27\r\nş. Prepeliţele şi mana, 16:1-36\r\nt. Stînca din Horeb, 17:1-7\r\nţ. Înfrîngerea lui Amalec, 18:1-12\r\nu. Judecătorii, 18:1-27\r\n\r\nII. LEGEA 19:1 - 24:18\r\na. Sinai, pregătiri pentru Legămînt, 19;\r\nb. Cele 10 porunci, 20:1-26\r\nc. Legi privitoare la robie, 21:1-11\r\nd. Legi pentru ucidere şi vătămare, 21:12-36\r\ne. Legi pentru furt, 22:1-4\r\nf. Legi pentru vătămarea averii, 22:5-15\r\ng. Legi împotriva imoralităţii, 22:16-20\r\nh. Legi pentru protejarea săracilor, 22:21-27\r\ni. Legi pentru îndatoriri religioase, 22:18-31\r\nî, Legi pentru îndatoriri civile, 23:1-14\r\nj. Lege pentru Sabat, 23:10-19\r\nk. Promisiune pentru cucerire, 23:20-33\r\nl. Ratificarea Legămîntului, 24:1-8\r\nm. Descoperirea slavei dumnezeieşti, 24:9-11\r\nn. Moise se suie pe munte, 24:12-18\r\n\r\nIII. CORTUL ÎNTÎLNIRII 25:1 - 40:38\r\n\r\nPLĂNUIT\r\na. Materialele pentru cort, 25:1-9\r\nb. Chivotul şi capacul ispăşirii, 25:10-22\r\nc. Masa cu pîinile, 25:23-30\r\nd. Sfeşnicul, 25:31-40\r\ne. Prelatele cortului, 26:1-14\r\nf. Scîndurile pentru pereţi, 26:15-30\r\ng. Cele două perdele, 26:31-37\r\nh. Altarul pentru arderile de tot, 27:1-8\r\ni. Curtea cortului, 27:9-19\r\nî. Untdelemnul sfînt, 27:20-21\r\nj. Îmbrăcămintea preoţească, 28:1-43\r\nk. Sfinţirea preoţilor, 29:1-46\r\nl. Altarul tămîierii, 30:1-10\r\nm. Darul pentru răscumpărare, 30:11-16\r\nn. Ligheanul de aramă, 30:17-21\r\no. Untdelemnul sfînt şi tămîia, 30:22-38\r\np. Rînduirea meşterilor, 31:1-11\r\nr. Sabatul ca semn, 31:12-17\r\ns. Tablele Legii, 31:18\r\n\r\nAMÎNAT\r\na. Viţelul de aur, 32:1-6\r\nb. Mînia lui Dumnezeu, 32:7-10\r\nc. Mijlocirea lui Moise, 32:11-14\r\nd. Mînia lui Moise, 32:15-29\r\ne. Moise se roagă pentru popor, 32:30-35\r\nf. Pocăinţa poporului, 33:1-11\r\ng. Moise cere să vadă slava, 33:12-23\r\nh. Reînnoirea Legămîntului, 34:1-35\r\n\r\nCONSTRUIT\r\na. Daruri pentru facerea conului, 35:1-29\r\nb. Chemarea meşterilor, 35:30-35\r\nc. Dărnicia poporului, 36:1-7\r\nd. Clădirea cortului, 36:8 - 38:31\r\ne. Facerea veşmintelor preoţeşti, 39:1-31\r\nf. Lucrarea este sfîrşită, 39:32-43\r\ng. Aşezarea şi sfinţirea cortului, 40:1-33\r\nh. Slava Domnului umple cortul, 40:34-38" },
   //LEVITICUL
    { "LEVITICUL", "Explicatia Cartii: Levitic\r\n\r\nCei mai mulţi dintre aceia care au avut dorinţa sinceră de a citi Biblia de la un capăt la altul s-au poticnit la cartea Leviticul! Trebuie s-o spunem şi noi: această carte nu este ca toate celelalte. Dar, faptul acesta nu o face mai puţin importantă, ci îi sporeşte şi mai mult însemnătatea unică. Deşi ea pare plicticoasă la citit, în ea se găsesc multe comori de extraordinară valoare. Plicticoasă pentru cititorul superficial, Leviticul se deschide ca o comoară de frumuseţi în faţa celui care este gata să facă un mic efort de studiu. Cartea aceasta nu ne-a fost dată pentru a fi citită, ci pentru a fi studiată.\r\n\r\nTitlul: Numirea din limba română vine din traducerea grecească cunoscută sub numele de „Septuaginta\" şi lămureşte specificul acestei cărţi. Leviticul este un cod de prevederi legale şi rituale dedicat Leviţilor, adică urmaşilor lui Levi, unul dintre cei 12 fii ai lui Iacov, care fuseseră aleşi de Dumnezeu să alcătuiască preoţia lui Israel (Gen. 46:1-27).\r\n\r\nAutorul: De 56 de ori este scris în text că Dumnezeu a vorbit aceste cuvinte lui Moise, care le-a scris apoi el însuşi sau le-a dictat altora spre a fi scrise (vezi Lev. 4:1; 6:1; 8:1; 11:1; 12:1). Domnul Isus certifică faptul că această carte îi aparţine lui Moise (Marcu l:44 cf. Lev. 13:49).\r\n\r\nContextul istoric: Cartea Exodul se termină cu terminarea construirii Cortului întîlnirii, făcut „după chipul care i s-a arătat lui Moise pe munte\". Cum trebuiau să folosească evreii această construcţie? Reglementările aşezate în cartea Leviticul sînt răspunsul la întrebarea aceasta. Ele i-au fost date de Dumnezeu lui Moise în timpul de o lună şi 20 de zile cît a durat construirea Cortului (Exod. 40:17) şi în perioada plecării poporului Israel de la muntele Sinai (Num. 10:11).\r\n\r\nConţinutul cărţii: Să spunem de la început că au existat patru plîngeri majore împotriva conţinutului acestei cărţi. Mai întîi au fost aceia care au susţinut că este imposibil să poţi înţelege şi explica ritualurile şi simbolurile cărţii Leviticul. Pentru aceşti critici, citirea cărţii nu poate aduce nici un fel de cîştig spiritual. Apoi sînt cei ce susţin că oricum cartea aceasta nu prezintă nici un fel de importanţă pentru creştinii de astăzi deoarece aşezămintele legămîntului Mozaic sînt depăşite şi n-au nici un fel de aplicabilitate în Biserică. A treia plîngere este aceia că impresia generală produsă de rigorile şi aşezămintele crude şi sîngeroase descrise în acest text sînt în contradicţie clară cu imaginea despre Dumnezeu pe care ne-o pun înainte celelalte cărţi ale Bibliei; în sfîrşit, a patra plîngere împotriva Leviticului vine din partea celor care zic că este foarte dificil de studiat această carte din cauză că nu este exprimată într-o expunere ordonată. Este uşor să urmăreşti desfăşurarea evenimentelor în Geneza şi Exodul, dar este aproape imposibil să înţelegi cum se succed lucrurile în Leviticul.\r\n\r\nUn foarte scurt cuvînt introductiv despre conţinutul Leviticului va risipi imediat aceste pretinse obstacole şi va arăta că această carte nu este depăşită de loc, că abundă în valori spirituale fără de care creştinismul veacului acesta ar fi mai sărăcit, că îl prezintă pe Dumnezeu în toată splendoarea caracterului Său şi că textul ei urmăreşte un plan prestabilit, armonios şi bine alcătuit. Toate cărţile care urmează în Biblie după Leviticul sînt marcate de culoarea specifică a acestei cărţi. Iată de ce este impoartant să o cunoaştem şi să ne familiarizăm cu terminologia şi cu conţinutul ei.\r\n\r\nCuvinte cheie şi teme caracteristice: Cartea aceasta este un manual al preoţimii şi este înţesat de termeni specifici ca aceştia: „jertfă\" (apare de 42 de ori), „preot\" (de 189 de ori), „sînge\" (de 86 de ori), „sfînt\" (de 87 de ori) şi „ispăşire\" (de 45 de ori). Reglementările din această carte subliniază necesitatea sfinţeniei în trup şi în suflet. Fraza: „să fiţi sfinţi, căci Eu sînt sfînt\" se repetă mereu ca un avertisment şi ca o aducere aminte (Lev. 11:44, 45; 19:2; 20:7, 26). Noul Testament citează din textul şi temele Leviticului de peste 90 de ori.\r\n\r\nScopul principal al cărţii: Leviticul a fost scris ca să arate Israelului cum să trăiască în sfinţenie. Lucrul acesta era necesar pentru că Dumnezeu hotărîse ca Israelul să fie instrumentul prin care să fie realizat planul de mîntuire pentru toate celelalte neamuri. Metoda lui Dumnezeu pentru imprimarea unui tipar de sfinţenie a fost întreită: jertfele, cu mesajul că „fără vărsare de sînge nu există iertare\", poruncile ca o mărturie a caracterului divin şi a conduitei cerute de Dumnezeu şi pedepsele, anexate poruncilor ca un avertisment despre inflexibilitatea sfinţeniei divine. Leviticul a fost destinat să pregătească Israelul şi omenirea pentru venirea lui Cristos, prin trezirea sentimentului de vinovăţie şi de nevoie după o jertfă de ispăşire, întregul ritual desfăşurat la Cortul întîlnirii era o trimitere spre lucrarea viitoare a Domnului Isus şi spre sacrificiul suprem depus de El pe altarul din dealul Golgotei.\r\n\r\nValoarea perpetuă: În primul rînd, Leviticul este o revelaţie a caracterului lui Dumnezeu, iar acesta este acelaşi ieri, astăzi şi în veci. În al doilea rînd, el prezintă felul lui Dumnezeu de a lucra cu oamenii. Nici acesta nu se schimbă. Sfinţenia lui Dumnezeu ne descopere mereu vinovăţia noastră şi nevoia noastră după Mijlocitorul, harului şi al iertării. În al treilea rînd, Leviticul este o mostră de cod moral din domeniul teocraţiei. Deşi textul este foarte vechi, orice grup legislativ care se respectă ar face bine să stăruiască asupra lui ori de cîte ori au de imprimat societăţii norme care să-i garanteze stabilitatea şi propăşirea. În al patrulea rînd, şi cel mai important dintre toate, Leviticul este o comoară de tipuri şi simboluri care ni-L prezintă pe Cristos în toată gloria şi splendoarea Lui trecută, prezentă şi viitoare. Nu trebuie să facem greşala să credem că dacă unele dintre aceste simboluri şi tipuri au fost deja împlinite, toate se găsesc în aceiaşi situaţie. Dacă este să ţinem seama de aluziile Noului Testament, multe tipuri şi simboluri îşi aşteaptă încă împlinirea. Iată de exemplu simbolurile înscrise în sărbătorile naţionale evreieşti. Primele două, Paştele şi Cincizecimea şi-au trăit împlinirea la cea dintîi venire a Domnului. Ce vom zice însă despre maiestuoasa Zi a Ispăşirii pentru întregul popor Israel? Am văzut din Evanghelii şi Epistole că Isus Cristos a intrat în cer ca Mare Preot şi a săvîrşit curăţirea tuturor lucrurilor, dar cum rămîne cu restul ritualului din sărbătoare: cu reîntoarcerea Lui la poporul care aştepta afară şi cu binecuvîntarea pe care trebuie s-o împartă asupra tuturor? S-a împlinit şi aceasta? Cum rămîne cu celelalte două sărbători din calendar? Ce se poate spune despre împlinirea Sărbătorii Trîmbiţelor şi despre acel magnific „An Sabatic\"? Dar cu extraordinarul an „al jubileelor\"? Istoria trecută şi prezentă nu prezintă nici un eveniment care să poată corespunde cu aceste semnificaţii. Prin urmare, Leviticul este dovedit a fi încă o carte care trebuie studiată pentru anticiparea, înţelegerea şi aşteptarea vremurilor viitoare. Acele vremi în care „pînă şi pe zurgălăii cailor va sta scris: „Sfinţi Domnului\" (Zaharia 14:20).\r\n\r\nPentru o înţelegere corectă a cărţii Leviticul trebuie să avem un punct de plecare corect. El se găseşte în primul verset din primul capitol al cărţii: „Domnul a chemat pe Moise; i-a vorbit din Cortul întîlnirii...”\r\n\r\nAnterior Dumnezeu le vorbise de pe muntele Sinai. Fusese un glas de Dumnezeu pentru un popor vinovat, cu care nu făcuse încă un legămînt. De data aceasta Dumnezeu vorbeşte poporului din interiorul Cortului. Este glasul Dumnezeului intrat într-o relaţie de părtăşie cu poporul Său prin intermediul jertfelor şi al iertării. Cartea Leviticul nu vorbeşte poporului ca să-l mîntuiască (lucrul acesta se săvîrşise prin jertfirea mielului pascal cu prilejul scoaterii din Egipt), ci pentru a-l învăţa cum să rămînă într-o relaţie bună cu Dumnezeul mîntuirii lor. În interpretările lor tipice, jertfele din această carte ne dezvăluie frumuseţea jertfei Domnului Isus în multiplele ei faţete prin care îşi exercită eficacitatea spre cei ce deja sînt mîntuiţi şi au intrat într-o relaţie nouă cu Dumnezeu fiind justificaţi prin credinţă. Acesta este punctul de plecare al cărţii Leviticul. El este o continuare firească a Genezei şi Exodului. În Geneza putem vedea remediul pregătit de Dumnezeu pentru ridicarea omului din căderea sa: „Sămînţa femeii\"; în Exodul vedem răspunsul lui Dumnezeu la strigătul desnădăjduit al omului aflat în robie: „sîngele Mielului\"; în Leviticul vedem tot ceea ce are nevoie omul ca să păstreze o relaţie bună cu Dumnezeu: un Preot, o Jertfă şi un Altar. Pe merit, Leviticul ocupă locul central în cele cinci cărţi ale lui Moise, căci prin învăţătura lui despre mijlocire prin preot, jertfă şi altar, el este adevărata inimă a Pentateucului - şi în fapt, a Evangheliei.\r\n\r\nLeviticul stă faţă de Exodul în aceiaşi relaţie pe care o au Epistolele cu Evangheliile, în Evanghelii noi sîntem „eliberaţi\" prin sîngele Mielului, în Epistole ajungem să fim locuiţi de Duhul Sfînt al lui Dumnezeu, în Evanghelii Dumnezeu ne vorbeşte din afară, în Epistole el glăsuieşte din lăuntru. În Evanghelii noi avem baza părtăşiei noastre cu Dumnezeu: răscumpărarea, în Epistole găsim umblarea noastră în părtăşia cu Dumnezeu: sfinţirea. Ca şi în Leviticul, şi în Epistole ne sînt prezentate multiplele faţete ale lucrării de ispăşire manifestate de Dumnezeu faţă de cei care sînt deja mîntuiţi.\r\n\r\nÎmpărţirea cărţii: Există două părţi distincte: pînă la capitolul 17 şi de la capitolul 17 la sfîrşit. Prima parte se ocupă de probleme ne-morale, iar cea de a două parte se ocupă de probleme morale. Prima parte are de a face cu întinăciunea ceremonială şi fizică, în timp ce a doua parte se ocupă de întinarea morală şi spirituală. Cea dintîi reglementează o curăţenie a inimii, iar cea de a doua defineşte limitele unei vieţuiri în absolută curăţie.\r\n\r\nSCHIŢA CĂRŢII\r\n\r\nI. BAZA PĂRTĂŞIEI: JERTFA 1 - 17\r\n\r\nJERTFELE\r\na. Arderile de tot, 1:1-17\r\nb. Jertfele de mîncare, 2:1-16\r\nc. Jertfele de mulţumire, 3:1-17\r\nd. Jertfele pentru ispăşire, 4:1-5:13\r\ne. Jertfele pentru vină, 5:14-6:7\r\nf. Lămuriri la legile expuse, 6:8-7:38\r\n\r\nPREOŢII\r\na. Consacrarea în slujbă, 8:1-36\r\nb. Inaugurarea slujirii, 9:1-24\r\nc. Denaturarea slujirii, 10:1-20\r\n\r\nOAMENII\r\na. Curăţia în mîncare, 11:1-47\r\nb. Curăţia după naştere, 12:1-8\r\nc. Curăţia în caz de lepra, 13:1-14:57\r\nd. Necurăţia bărbatului, 15:1-18\r\ne. Necurăţia femeii, 15:19-33\r\nf. Sărbătoarea ispăşirii, 16:1-34\r\n\r\nALTARUL\r\na. Singurul loc pentru jertfă, 17:1-9\r\nb. Interdicţia de a mînca sînge, 17:10-16\r\n\r\nII. ROADA PĂRTĂŞIEI: SFINŢIREA 18-27\r\n\r\nOAMENII\r\na. Sfinţire în relaţiile trupeşti, 18:1-30\r\nb. Sfinţire în viaţa zilnică, 19:1-37\r\nc. Pedepse pentru sfidarea sfinţirii, 20:1-27\r\n\r\nPREOŢII\r\na. Porunci privitoare la preoţi, 21:1-22:16\r\n\r\nSĂRBĂTORILE\r\na. Sabatul, 23:1-3\r\nb. Paştele, 23:4-8\r\nc. Rusaliile, 23:8-22\r\nd. Anul nou şi trîmbiţele, 23:23-25\r\ne. Ziua ispăşirii, 23:26-32\r\nf. Sărbătoarea corturilor, 23:33-44\r\n\r\nULEIUL Şl PÎINEA\r\na. Untdelemnul pentru sfeşnic, 24:1 ;4\r\nb. Pîinile pentru punerea înainte, 24:5-9\r\n\r\nFELURITE PEDEPSE\r\na. Pedeapsa celui ce huleşte, 24:10-16\r\nb. Pedeapsa celui ce ucide, 24:17-23\r\n\r\nRÎNDUIELI PENTRU ŢARA CANAANULUI\r\na. Odihna pămîntului, 25:1-7\r\nb. Anul de veselie, 25:8-55\r\nc. Binecuvîntarea şi blestemul, 26:1-46\r\nd. Pocăinţa, 26:40-46\r\ne. Juruinţele, 27:1-29\r\nf. Zeciuielile, 27:30-34" },
    //NUMERI
    { "NUMERI", "Explicatia Cartii: Numeri\r\n\r\nÎn mod prioritar, Vechiul Testament este cartea poporului evreu. Toate cărţile lui trebuiesc judecate prin prisma existenţei şi destinului acestui popor de excepţie. Cartea Genezei ne-a arătat coordonatele cosmice şi terestre în care a apărut acest neam. Capitolul 11 din cartea Genezei este pragul de debut al poporului Israel. Dumnezeu însuşi intervine în istorie şi-l cheamă pe Avram din cetatea păgînă Ur, aflată în Caldeea, la vărsarea Tigrului şi Eufratului în golful Persic. Ales de gestul suveran al lui Dumnezeu, acest Avram devine Avraam, tatăl unei mari mulţimi (Gen. 17:5), şi strămoşul tuturor evreilor. Finalul cărţii Geneza îi găseşte pe urmaşii acestui foarte important personaj istoric părăsind temporar Canaanul spre a se muta în Egipt.\r\n\r\nCartea Exodul ne arată cum familia celor 12 seminţii ale lui Iacov, nepotul lui Avraam creşte numeric depăşind stadiul unui simplu grup etnic şi devine un popor care stîrneşte teama egiptenilor. Citim apoi cum Dumnezeu însuşi intervine cu mînă tare în Egipt spre a-l determina pe Faraon să-i lase pe evrei să plece spre a se întoarce înapoi în ţara promisă lui Avraam, ţara Canaanului. După ieşirea din Egipt, Dumnezeu dă acestui popor un cod de legi teocratice în virtutea căruia „Elohim\" Dumnezeul cel atotputernic, devine Iehova, Dumnezeul evreilor. Israelul îşi întăreşte statutul de popor ales şi îşi însuşeşte o viaţă cultică corespunzătoare. Leviticul redă amănunţit prevederile morale, sociale şi ceremoniale ale codului teocraţie de legi date evreilor. Cartea Numeri, la care am ajuns acum, cuprinde încă o etapă în drumul împlinirii fiinţei naţionale evreieşti. Poporul format în Egipt şi căruia i s-a dat în pustie o Constituţie, este acum în drum spre vatra în care-şi va aşeza copiii. Dumnezeu îi conduce înainte ca să le dea Canaanul. Se vor lăsa ei conduşi de Dumnezeu în această ţară nouă şi prosperă, dar necunoscută şi potenţial pericuIoasă? Aceasta este întrebarea la care răspunde conţinutul acestei cărţi.\r\n\r\nTitlul: Evreieşte, cartea se numeşte: „În pustie\". Cei 70 de învăţaţi care au tradus Biblia din ebraică în greceşte la Alexandria în Egipt, au numit-o în Septuaginta: „Aritmoi\", numărătorile. Titlul din limba română vine din traducerea latină în care această carte se numeşte chiar aşa: „Numeri\".\r\n\r\nAutorul: Toată lumea este unanimă în părerea că autorul cărţii este Moise.\r\n\r\nData scrierii: Probabil între anii 1450-1410 î.Cr.\r\n\r\nContextul istoric: La o lună şi jumătate după clădirea Cortului are loc cea dintîi numărătoare a poporului (Numeri 1:1- 4:49). Acest recensămînt a fost necesar pentru a determina capacitatea militară a naţiunii şi pentru a rîndui aşezarea taberei lui Israel în jurul clădirii Cortului. Fără nici o îndoială că numărătoarea a fost necesară şi pentru reglementările administrative şi economice. Mulţimea de peste două milioane de persoane avea nevoie de puţină ordine şi disciplină economică şi administrativă. Apoi urmează marea tragedie: poporul refuză să se încreadă în Dumnezeu pentru cucerirea Canaanului şi Dumnezeu hotărăşte să-i rătăcească prin pustie timp de 40 de ani. După această perioadă are loc cea de a doua numărătoare (Num. 26:1- 65). Toată mulţimea celor ce trecuseră de 20 de ani cu prilejul neascultării, pieriseră deja în pustie. Numai Iosua şi Caleb au fost cruţaţi. Ei vor conduce de acuma poporul. Au loc ultimele pregătiri înainte de a trece la cucerirea ţării. Poporul adunat la Cadeş este numărat încă o dată. Numărul lor este aproximativ egal cu numărul celor ce, cu 40 de ani în urmă, fuseseră scoşi din Egipt.\r\n\r\nTema cărţii: Cineva a zis că: (1) în Geneza vedem ruina omului, (2) în Exodul vedem răscumpărarea omului, (3) în Levitic vedem închinăciunea omului, iar (4) în Numeri vedem omul în umblarea lui cu Domnul. Şi numărătorile şi pribegia timp de 40 de ani prin pustie susţin acest punct de vedere.\r\n\r\nLecţia principală din cartea Numeri este că poporul lui Dumnezeu trebuie să umble prin credinţă dacă vrea să înainteze. Pentru a sublinia acest adevăr, textul relatează abateri de la această regulă şi sfîrşitul lor tragic: neascultarea întregului popor (Num. 11:1), neascultarea lui Aaron şi Maria (Num. 12:1, 9-10), refuzul de intra în ţara promisă (Num. 14:2-3, 26-30), neascultarea lui Moise (Num. 20:12), idolatria poporului (Num. 25:3, 8-9). În ciuda acestor greşeli ale poporului, Dumnezeul legămîntului făcut cu ei i-a purtat cu credincioşie avînd grijă în mod miraculos de ei în cei 40 de ani de pribegie în pustie şi i-a făcut să ajungă în final în binecuvîntată ţară a Canaanului.\r\n\r\nNoul Testament preia lecţia neascultării din cartea Numeri şi o foloseşte pentru a imprima creştinilor o viaţă de credinţă şi de deplină ascultare faţă de Domnul (Ioan 3:14; l Cor. 10:1-12; 2 Petru 2:l5-16; Apoc. 2:14; Iuda 11)\r\n\r\nConţinutul cărţii: Aşezarea conţinutului acestei cărţi este unică în felul ei şi, o dată înţeleasă, ne va ajuta să păstrăm foarte uşor în memorie evenimentele cărţii. De la o primă privire observăm că avem de a face cu două generaţii de oameni: cea dintîi, care ieşise din Egipt dar care avea să piară în pustie, şi cea de a doua, care s-a născut în pustie şi avea să cucerească Canaanul. Primele 14 capitole tratează problema primei generaţii, ultimele 16 capitole istorisesc ceea ce s-a întîmplat cu generaţia născută în pustie, iar între ele există grupul capitolelor 15-20 în care aflăm ceea ce s-a întîmplat pe drumurile pustiei în acea perioadă de tranziţie de 40 de ani. De unde ştim că de la capitolul 21 începe relatarea evenimentelor din viaţa generaţiei născute în pustie? Textul din Num. 33:38 ne spune că moartea lui Iosua descrisă în Num. 20:22-29 s-a petrecut: „în al patruzecilea an după ieşirea copiilor lui Israel din ţara Egiptului, în luna a cincea, în cea dintîi zi a lunii\".\r\n\r\nCuvinte cheie şi teme caracteristice: Cartea Numeri mai poate fi numită şi „cartea călătoriilor şi popasurilor\" (Num. 33:1-2) sau „cartea cîrtirilor\" (Num. 11:1, 4; 12:2; etc. Ps. 95:10). Cel mai cunoscut pasaj al cărţii este întîmplarea cu şarpele de aramă (Num. 21:4-9) pe care o preia Domnul Isus în Ioan 3:14-17 aplicînd-o la propria Lui răstignire pentru mîntuirea lumii. Alte tipuri care îl prevestesc pe Cristos sînt: stînca lovită (Num. 20:7-11) şi cetăţile de scăpare (Num. 35).\r\n\r\nInsuficienţa aşezămintelor Vechi Testamentale pentru mîntuire este ilustrată profetic de evenimentele petrecute în Num. 20. În acest unic capitol ne întîlnim cu moartea Mariei, cu păcatul lui Moise şi cu moartea lui Aaron. Maria este reprezentativă pentru oficiul profeţiei, Aaron este reprezentantul preoţiei Vechi Testamentale, iar Moise este prin excelenţă reprezentantul Legii. Nici unul dintre aceşti trei oameni nu a putut intra în ţara promisă. Această misiune a fost păstrată pentru Iosua, care în nume şi misiune este un reprezentant al Căpeteniei mîntuirii noastre: Domnul Isus Cristos.\r\n\r\nSCHIŢA CĂRŢII\r\n\r\nI. CEI IEŞIŢI DIN EGIPT 1-14\r\na. Numărarea Israeliţilor, 1:1-4:49\r\nb. Poziţia semintiilor în tabără, 2:1-34\r\nc. Poziţia şi slujbele leviţilor, 3:1-4:49\r\n\r\nSFINŢIREA POPORULUI\r\na. Prin izgonirea celor necuraţi, 5:1-4\r\nb. Prin despăgubirea pentru furt, 5:4-10\r\nc. Prin clarificări în căsnicie, 5:11-31\r\nd. Prin legea Nazireatului, 6:1-21\r\ne. Prin binecuvîntarea poporului, 6:22-27\r\nf. Prin darurile căpeteniilor, 7:1-89\r\ng. Prin închinarea leviţilor, 8:1-26\r\nh. Prin ţinerea Paştelor, 9:1-14\r\ni. Prin călăuzirea divină, 9:15-10:10 P\r\n\r\nPLECAREA DIN SINAI\r\na. Începutul călătoriei, 10:11-28\r\nb. Socrul lui Moise, 10:29-36\r\nc. Începutul cîrtirilor, 11:1-3\r\nd. Cîrtirea poporului, 11:4-35\r\ne. Cîrtirea Mariei, 12:1-16\r\n\r\nISCODIREA ŢĂRII\r\na. Trimiterea celor 12 iscoade, 13:1-20\r\nb. Iscodirea ţării, 13:21-25\r\nc. Raportul iscoadelor, 13:26-33\r\nd. Necredinţa poporului, 14:1-10\r\ne. Mînia lui Dumnezeu, 14:11-12\r\nf. Mijlocirea lui Moise, 14:13-19\r\ng. Pedepsirea celor necredincioşi, 14:20-38\r\nh. Pedepsirea celor neascultători, 14:39-45\r\n\r\nII. 40 DE ANI ÎN PUSTIE 15-20\r\na. Legi asupra jertfelor, 15:1-21\r\nb. Păcatele fără voie, 15:22-31\r\nc. Călcarea Sabatului, 15:32-36\r\nd. Ciucurii de la veşminte, 15:37-41\r\n\r\nCONTESTAREA Şl CONFIRMAREA PREOŢIEI\r\na. Core, Datan şi Abiram, 16:1-50\r\nb. Toiagul lui Aaron, 17:1-13\r\nc. Responsabilităţile preoţilor, 18:1-7\r\nd. Privilegiile preoţilor, 18:8-32\r\ne. Vaca roşie, apa de curăţie, 19:1-22\r\n\r\nMARIA, MOISE, AARON\r\na. Moartea Mariei, 20:1\r\nb. Păcatul lui Moise, 20:2-13\r\nc. Împotrivirea Edomiţilor, 20:14-21\r\nd. Moartea lui Aaron, Eleazar, 20:22-29\r\n\r\nIII. CEI NĂSCUŢI ÎN PUSTIE 21-36\r\na. Înfrîngerea şi biruinţa de la Arad, 21:1-3\r\nb. Şarpele de aramă, 21:4-9\r\nc. Alte călătorii, 22:10-35\r\n\r\nBALAAM, PROFETUL VRĂJITOR\r\na. Balaac trimite după Balaam, 22:1-14\r\nb. Măgăriţa lui Balaam, 22:15-41\r\nc. Balaam binecuvîntează, 23:1-30\r\nd. Balaam confirmă binecuvîntarea, 24;1-14\r\ne. Steaua lui Iacov, 24:15-25\r\nf. Şiretlicul lui Balaam, 25:1-15 (31:16)\r\ng. Omorîrea Madianiţilor, 25:16-18\r\n\r\nO NOUĂ NUMĂRĂTOARE\r\na. A doua numărătoare, 26:1-65\r\n\r\nREPETAREA UNOR LEGI\r\na. Lege asupra moştenirii, 27:1-11\r\nb. Iosua, urmaşul lui Moise, 27:12-23\r\nc. Jertfa zilnică, 28:1-15\r\nd. Jertfe de Paşte, 28:16-25\r\ne. Jertfele de Rusalii, 28:26-31\r\nf. Jerfe în ziua trîmbiţei, 29:1 -6\r\ng. Jertfe pentru ziua ispăşirii, 29:7-11\r\nh. Jertfe pt. sărbătoarea corturilor, 29:12-40\r\ni. Lege asupra juruinţelor, 30:1-2\r\nî. Juruinţa femeilor, 30:3-16\r\n\r\nBIRUINŢA ASUPRA MADIANIŢILOR\r\na. Lupta, 31:1-24\r\nb. Împărţirea prăzii, 31:25-54\r\n\r\nŢARA DE DINCOLO DE IORDAN\r\na. Ruben şi Gad aleg Galaadul, 32:1-32\r\nb. Cucerirea Galaadului, 32:33-42\r\n\r\nRECAPITUALREA CĂLĂTORIILOR\r\na. Din Egipt pînă la Sinai, 33:15\r\nb. De la Sinai pînă la Cades, 33:16-36\r\nc. De la Cades la cîmpia Moabului, 33:37-49\r\n\r\nVOIA DOMNULUI PENTRU CANAAN\r\na. Canaaniţi trebuiesc nimiciţi, 33:50-56\r\nb. Hotarele Canaanului, 34:1-29\r\nc. Cetăţile Leviţilor, 35:1-8\r\nd. Cetăţile de scăpare, 35:9-34\r\ne. Moştenirea femeilor, 36:1-13" },
   //DEUTERONOMUL
    { "DEUTERONOMUL", "Explicatia Cartii: Deuteronom \r\n\r\nDeuteronomul este cea de a cincea carte din Pentateucul lui Moise. Aceste cinci carti, numite de evrei si Tora, alcatuiesc o veritabila Biblie în miniatura, în ele vedem succesiv, ruina, rascumpararea si readucerea omului în partasie cu Dumnezeu, apoi calauzirea divina si dragostea lui Dumnezeu care reuseste mereu sa ne scoata la capat. ªi tot în ele capatam o revelatie progresiva treptata a lui Dumnezeu catre oameni. Geneza ne descopere suveranitatea lui Dumnezeu, în creatie si în alegerea Israelului; Exodul ne descopere puterea si maretia lui Dumnezeu; Leviticul adauga la acestea sfintenia divina, evidentiata în îndemnul la punerea de o parte pentru Domnul; în Numeri ni se arata severitatea si bunatatea manifestata de Dumnezeu în relatia Lui cu oamenii, iar în Deuteronomul vom vedea dincolo de promisiuni si pedepse credinciosia lui Dumnezeu. Aceasta carte devine astfel nu numai ultima dintre cele cinci, ci si o completare necesara a tuturor celorlalte. Numele cartii vine de la doua cuvinte din limba greaca: deuteros (a doua) si nomos (lege). În secolul III dinainte de Cristos, învatatii care au tradus originalul ebraic în Septuaginta greaca au combinat aceste cuvinte pentru a obtine titlul cartii, în fapt, Deuteronomul nu este o a doua lege, ci numai o repetare a legilor înaintea noii generatii de evrei nascuti în pustie si aflati acum în pragul intrarii în Canaan. Asa cum arata si continutul, cu exceptia ultimului capitol în care moartea lui Moise este adaugata de altcineva ca un fel de apendix al cartii, Deuteronomul este duiosul cîntec de lebada al lui Moise. Adevaratul autor al cartilor sfinte este însa altcineva. De peste 500 de ori ni se spune în textul Pentateucului ca: „Dumnezeu a spus:...” sau „Domnul a vorbit lui Moise si i-a spus:...” Desi vom da mereu în schitele noastre numele autorului uman al cartilor Bibliei este bine sa tinem minte ca veritabilul autor al lor ramîne Dumnezeu însusi. Iata ce ne spune Petru în aceasta privinta: „Caci nici o proorocie n-a fost adusa prin voia omului; ci oamenii au vorbit de la Dumnezeu, mînati de Duhul Sfînt\" (2 Petru l:21; vezi si 2 Tim. 3:16, 17). Moise rosteste aceste cuvinte în cîmpia Moabului, înainte de intrarea evreilor în Canaan si cu putin înainte de propria lui moarte. Deuteronomul acopera o perioada de timp de aproximativ doua luni, incluzînd si cele 30 de zile de jale dupa moartea lui Moise. Probabil în timpul unei singure saptamîni, cam cu o luna înainte de trecerea Iordanului (Deut. 1:1-3) Cartea este o colectie de cuvîntari tinute de Moise ca ramas bun înainte de despartirea de copiii lui Israel. Aceste discursuri calde si solemne au fost rostite în ceasul în care marele om se afla în pragul marii treceri din viata terestra spre viata cereasca. Din înaltimea muntelui Pisga, el priveste înapoi peste un secol de istorie plin cu evenimente de importanta epocala. Dupa înca o sectiune în care revine cu lamuriri la unele din legile primite anterior de Israel, Moise îsi întoarce privirea înspre viitor patrunzînd profetic pîna catre sfîrsitul istoriei lui Israel, dupa anticipata lor reîntoarcere în tara.\r\nCuvinte cheie si teme caracteristice: Deuteronomul este o carte de tranzitie. Ea marcheaza patru schimbari mari care s-au petrecut în viata Israelului. În primul rînd este vorba de tranzitia înspre o noua generatie. Cu exceptia lui Iosua, Caleb si a lui Moise însusi, toata multimea aceea de oameni scoasa din Egipt si numarata la Sinai se prapadise pe drumurile pustiei, în al doilea rînd este vorba despre o tranzitie înspre o noua mostenire. Peregrinarile prin pustie aveau sa se sfârseasca prin dobîndirea unui teritoriu în care sa se aseze natiunea. În cel de al treilea rînd este vorba despre tranzitia spre o noua experienta, spre un alt fel de viata în care corturile vor fi înlocuite cu case, viata nomada va fi parasita si în loc de hrana pustiei vor începe sa manînce lapte, miere si roadele Canaanului. Cea de a patra tranzitie, si cea mai importanta va fi aceea înspre o noua cunoastere a lui Dumnezeu, înspre cunoasterea ca Dumnezeu este dragoste.\r\nÎncepînd cu Geneza si pîna la Numeri, textul Scripturii nu ne vorbeste nicaieri explicit despre dragostea divina. Abia în Deuteronomul, li se explica evreilor motivatia faptelor lui Dumnezeu: „El a iubit pe parintii tai, si de aceea a ales samînta lor dupa ei\" (Deut. 4:37). „Nu doar pentru ca întreceri la numar pe toate celelalte popoare S-a alipit Domnul de voi si v-a ales, caci voi sînteti cel mai mic dintre toate popoarele. Ci, pentru ca Domnul va iubeste...” (Deut. 7:7-8). „ªi numai de parintii tai S-a alipit Domnul ca sa-i iubeasca; si dupa ei, pe samînta lor, pe care v-a ales El dintre toate popoarele, cum vedeti azi\" (Deut. 10:5; vezi si 23:5).\r\nVorbind de acest caracter de tranzitie al cartii Deuteronomul, este bine sa aratam aici simetria care exista între Vechiul Testament si Noul Testament. Ambele încep cu cîte un grup de cinci carti istorico-legislative si exista o asemanare uimitoare între cea de a cincea carte a Vechiului Testament: Deuteronomul si cea de a cincea carte a Noului Testament: Faptele Apostolilor. Cartea Faptelor Apostolilor este si ea o carte de tranzitie. Ea marcheaza trecerea de la mesajul Evangheliilor la acela al epistolelor. Ca si Deuteronomul, si Faptele Apostolilor se adreseaza unei generatii noi: generatia „re-generatilor\" prin credinta mîntuitoare în Cristos. Ca si Deuteronomul si ea vorbeste despre o noua posesiune: „Canaanul spiritual plin de tot felul de binecuvîntari ceresti în Cristos\". Ca si Deuteronomul si ea vorbeste despre o noua experienta: experienta unei nasteri noi, a unei vieti noi, a unei puteri noi prin locuirea Duhului Sfînt în cei credinciosi. Ca si Deuteronomul si ea marcheaza tranzitia înspre o noua revelatie a lui Dumnezeu: „Taina lui Cristos... care n-a fost facuta cunoscut fiilor oamenilor în celelalte veacuri... pentru ca domniile si stapînirile din locurile ceresti sa cunoasca acum, prin Biserica, întelepciunea nespus de felurita a lui Dumnezeu\" (Efes. 3:4, 5, 10).\r\nCeea ce este înca si mai izbitor este faptul ca si Deuteronomul si Faptele Apostolilor sînt cartile în care Dumnezeu le da oamenilor o a doua sansa. Care sansa este în Deuteronomul? Este tocmai aceasta „deuteros nomos\", aceasta repetare a legii, înainte de a-i preda conducerea lui Iosua, Moise le repeta evreilor legamîntul Legii. Despre care a doua sansa se vorbeste în cartea Faptele Apostolilor? Este vorba despre cea de a doua oferire a împaratiei cerurilor pentru poporul evreu, mai întîi în capitala, iar apoi în tara si în toate provinciile imperiului Roman. Despre aceste lucruri vom vorbi mai pe larg atunci cînd vom ajunge la ele, dar este bine sa sa le cunoastem înca de pe acum.\r\nDeuteronomul vorbeste si despre trinitatea lui Dumnezeu. Iata ce citim în pasajul din 6:4-5, numit de evrei si „Shema\": „Asculta, Israele! Domnul, Dumnezeul nostru este un singur Domn. Sa iubesti pe Domnul, Dumnezeul tau, cu toata inima ta, cu tot sufletul tau si cu toata puterea ta\". Domnul Isus însusi ne-a spus ca aceasta este cea mai mare dintre toate afirmatiile Legii (Marcu 12:29-30). Evreii si unitarienii vad în acest text o negare a trinitatii, dar o asemenea afirmatie intra în conflict direct cu afirmatiile textului. Cuvîntul folosit pentru numele lui Dumnezeu este aici: „Elohenu\" forma genitivala a lui Elohim, despre care stim deja ca este o forma de plural. Citit în original, acest verset suna: „Iehova Elohenu, este singurul Iehova!\" sau tradus: „Domnul Dumnezeii nostri, este singurul Domn!\" Ba înca si mai mult, pentru cuvîntul: „singurul\" este folosit „echad\" care numeste o pluralitate si nu „iachid\" care defineste unitatea absoluta. Singularul de pluralitate se foloseste si în limba româna atunci cînd zicem de exemplu: „un\" ciorchine de struguri.\r\nÎn aceasta declaratie solemna din Legea repetata de Moise ni se da cea mai înalta afirmatie de crez a religiei poporului Israel. Ea este astazi si temelia crezului crestin. Nici nu se putea sa fie altfel. Dumnezeul evreilor este si Dumnezeul nostru, caci în afara Lui nici nu exista un altul. Mesia evreilor este Mesia al nostru, caci nici aici nu exista un altul. Crestinismul este monoteist si cristocentric.\r\nUn alt pasaj de importanta exceptionala din Deuteronom este capitolul 28. El ne arata ceea ce ar fi putut deveni Israel daca ar fi trait în ascultare de Domnul (28:1-14). Cum asa ceva nu s-a întîmplat, pasajul ramîne o referinta despre vremea „mileniului\" (vezi si Isaia 60 pîna la 62; Zaharia 14:8-21, Ieremia 31:1-19, Deut. 30:1-10; Rom. 11:25-31). Versetele 47-49 ne vorbesc despre invazia romana din anul 70 d.Cr. si de cumplita tragedie a macelaririi evreilor, iar versetele 63-67 ni-i descriu pe evrei asa cum traiesc ei astazi, „împrastiati\" printre toate popoarele lumii, „fara loc de odihna\", cu „inima fricoasa\", asteptînd îndurarea Dumnezeului lor si venirea lui Mesia.\r\nMoise a murit la vîrsta de 120 de ani (Deut. 34:7). El este singurul om pe care l-a îngropat Dumnezeu însusi. Mormîntul lui a fost asezat „în valea Moabului\", dar nimeni nu i-a cunoscut locul „pîna în ziua de astazi\" (34:6). Dumnezeu a evitat astfel pericolul ca evreii sa faca din Moise un idol. Dumnezeul crestin s-a pronuntat înca de atunci împotriva cultului mortilor. Biblia ne spune ca pentru trupul lui Moise a fost o disputa între Diavol si arhanghelul Mihail (Iuda 9) Moise a reaparut înca o data pe pamînt pe muntele unde Domnul S-a schimbat la fata înaintea ucenicilor Sai (Luca 9:30, 31). El a trait dincolo de mormînt, continuînd sa-L serveasca pe Dumnezeul rascumpararii sale.\r\nSCHItA CaRtII\r\nIntroducere 1:1-5\r\nI. O PRIVIRE ÎN URMa 1:6 - 4:33\r\na. Israel la Sinai, 1:6-18\r\nb. Israel la Cades-Barnea, 1:19-46\r\nc. Pribegia prin pustie, 2:1-25\r\nd. Cucerirea tarii dinainte de Iordan, 2:26-3:22\r\ne. Moise cere sa vada tara, 3:23-29\r\nf. Îndemn, 4:1-40\r\ng. Cetatile de scapare, 4:41-43\r\nII. O PRIVIRE ÎN SUS 4:44 - 26:19\r\na. Cele 10 porunci, 4:44-5:33\r\nb. Iubirea de Dumnezeu, 6:1-25\r\nc. Nimicirea Cananitilor, 7:1-26\r\nd. Porunca aducerii aminte, 8:1-10:11\r\ne. Ce cere Dumnezeu, 10:12-11:25\r\nf. Binecuvântarea si blestemul, 11:26-32\r\ng. Stîrpirea idolatriei, 12:1-13:18\r\nh. Mîncari curate si necurate, 14:1-21\r\ni. Zeciuielile, 14:22-29\r\nî. Anul sabatic, 15:1-23\r\nj. Sarbatorile evreiesti, 16:1-17\r\nPorunci pentru:\r\nk. judecatori, 16:18-17:13\r\nl. împarati, 17:14-20\r\nm. Leviti, 18:1-18\r\nn. ghicitori si vrajitori, 18:9-14\r\no. primirea lui Mesia, 18:15-22\r\np. profeti, 18:20-22\r\nPorunci pentru relatii sociale:\r\nr. Cetatile de scapare, 19:1-13\r\ns. Hotarele proprietatii, 19:14\r\ns. Martorii mincinosi, 14:15-21\r\nt. Militarie, 20:1-14\r\nt. Cetatile cucerite, 20:15-20\r\nu. Omoruri, 21:1-9\r\nv. Viata de familie, 21:10-22:30\r\nx. Prozelitii si adunarea Domnului, 23:1-18\r\ny. Protejarea celui mai slab, 23:19-25:19\r\nz. Cele dintîi roade, 26:1-19\r\nIII. O PRIVIRE ÎNAINTE 27 - 30\r\na. Stabilirea cadrului, 27:1-26\r\nb. Binecuvîntarile, 28:1-14\r\nc. Blestemurile, 28:15-68\r\nd. Legamîntul Palestinian, 29:1-30:20\r\nÎncheiere\r\na. Cuvinte de ramas bun, 31:1-29\r\nb. Cîntarea lui Moise, 31:30-32:47\r\nc. Testamentul lui Moise, 32:48-33:29\r\nd. Moartea lui Moise, 34:1-12\r\nCauta in biblie\r\nTextul cautat\r\nCauta...\r\nCum caut\r\nToate Cuvintele\r\nUnde caut:\r\nToată biblia\r\nVersetul Zilei\r\nDaniel 5:26-28\r\n\r\nIata insa scrierea care a fost scrisa: Mene - DUMNEZEU ti-a numarat zilele domniei, si i-a pus capat. Techel - inseamna ca ai fost cantarit in cumpana si ai fost gasit usor. Peres - inseamna ca imparatia ta va fi impartita, si data Mezilor si Persilor.\r\nMeditatia Zilei\r\n\r\nPowered by Biblia Online\r\nFacebook\r\n\r\nParteneri\r\nStiri Crestine\r\nMuzica Crestina\r\nVersuri Crestine\r\nChristian Lyrics\r\nBiblia Online Cornilescu\r\nVersuri Crestine\r\nNu este Crăciun fără Isus - Estera & Laura Bretan\r\nÎn iesle azi - Diana Scridon Pop\r\nBucurie in suflet - Familia Timofte\r\nSe-aude Glas Peste Carpați / Tăria Noastră Fii Mereu - The Beuca Family\r\nUn singur Dumnezeu - Otto Pascal & Biji\r\nFiu iubit - BBSO\r\nClipa de clipa - Philadelphia Band\r\nPrintul Pacii a venit - Alin și Emima Timofte & TB Music Family\r\nPraise - Lumina Worship\r\nUnde ești țara mea?\r\nPowered by VersuriCrestine.ro\r\n" },
    //IOSUA
    { "IOSUA", "Explicatia Cartii: Iosua\r\n\r\n„...mergeti sa cuceriti tara pe care v-o da în stapînire Domnul, Dumnezeul vostru\" (Iosua 1:11). Este bucuria lui Dumnezeu sa daruiasca, este datoria noastra sa cucerim prin credinta!\r\nPentru ostasul crucii nu exista nici o alta carte mai plina de încurajare si de învatatura ca aceasta cronica a viteazului Iosua. Paginile ei mustesc de adevaruri spirituale valabile pentru copiii lui Dumnezeu din toate timpurile.\r\nAceasta carte este prima carte din Biblie care poarta numele autorului ei. De fapt, numele initial al acestui om a fost „Hoseea\" sau poate „Iehosua\" care se pot traduce prin perifraza: „Dumnezeu este Mîntuitorul\". Pus în slujba Domnului înca de tînar, omul acesta a devenit: „Iosua\" care tradus înseamna: „Robul Domnului\" sau „Cel ce împlineste lucrarile Domnului\".\r\nEste foarte clar ca Iosua s-a nascut în Egipt si se prea poate sa fi servit chiar în armata egipteana. În orice caz, îl gasim foarte priceput în tactica militara atunci cînd este asezat în fruntea ostilor lui Israel sa duca lupta „din vale\" împotriva amalecitilor la Refidim (Exod 17:8-16). Iosua a fost aghiotantul lui Moise în timpul evenimentelor petrecute la muntele Sinai (Exod 24:13), iar ca reprezentant al tribului lui Efraim a facut parte din grupul celor 12 iscoade care au fost trimise sa cerceteze tara Canaanului (Num 13:1-33). Iosua si Caleb au fost singurii dintre cei 12 care au încurajat poporul sa porneasca deîndata la cucerirea tarii (Num. 14:6-9)\r\n1.400-1.370 î.Cr. Cartea acopere o perioada de aproximativ 25 de ani din istoria lui Israel.\r\nCartea Iosua continua actiunea de la sfîrsitul cartii Deuteronomul. Dupa moartea lui Moise, Iosua devine noul conducator al poporului si sub conducerea lui evreii duc campaniile de cucerire a Canaanului si tot sub conducerea lui împart tara între cele 12 semintii ale lui Israel.\r\n\r\nCuvinte cheie si teme caracteristice: Întreaga carte este un tip cu semnificatie spirituala. Talmacirea acestui tip este data de Noul Testament în Evrei capitolele 3 si 4. Aceste doua capitole trebuiesc citite cu cea mai mare atentie. Trecerea Iordanului si cucerirea tarii este echivalata aici cu intrarea în prerogativele unei vieti spirituale binecuvîntate prin unirea cu Cristos Isus nu dupa moartea noastra, ci înca din aceasta viata. Canaanul este simbolul vietii crestine abundente care este daruita celor ce stiu sa avanseze prin credinta. C.H. Spurgeon spunea: „Exista o viata de plinatate crestina care este la fel de deosebita de viata crestina obisnuita, precum este aceasta de deosebita de viata de saracie spirituala a lumii\". Canaanul le-a dat atunci evreilor: (1) odihna, (2) abundenta, (3) biruinta. Asa ceva este pregatit si pentru cei ce se apropie cu credinta de Domnul Isus: „Ramîne dar o odihna ca cea de Sabat pentru poporul lui Dumnezeu. Fiindca cine intra în odihna Lui se odihneste si el de lucrarile lui, cum S-a odihnit Dumnezeu de lucrarile Sale.\" (Evrei 4:8-11) „...noi, fiindca am crezut, intram în „odihna\", despre care a vorbit El...” (Evrei 4:3)\r\n\r\nDa, Canaanul a trebuit cucerit (Iosua 1:1-2), dar evreii trebuiau sa faca aceasta nu prin forta lor, ci prin credinta (Deut. 7:1; Deut. 6:10-11; Lev. 26:6; Deut. 11:10-12). Forta nu le-a folosit la nimic cînd pacatul i-a facut sa fie învinsi la Ai, dar credinta si ascultarea au darîmat zidurile Ierihonului (Iosua 8; Iosua 6).\r\n\r\nSCHItA CaRtII\r\n\r\nI. INTRAREA ÎN CANAAN 1:1-5:12\r\na. Dumnezeu fi vorbeste lui Iosua, 1:1-9\r\nb. Iosua vorbeste poporului, 1:10-15\r\nc. Poporul fagaduieste ascultare, 1:16-18\r\nd. Iscodirea Ierihonului, Rahav, 2:1-24\r\ne. Trecerea tordanului, 3:1-17\r\nf. Pietre de aducere aminte, 4:1-24\r\ng. Taierea împrejur, 5:1-8\r\nh. Roadele tarii, mana înceteaza, 5:9-12\r\n\r\nII. CUCERIREA taRII 5:13-12:24\r\na. Capetenia ostirii Domnului, 5:13-15\r\nb. Cucerirea Ierihonului, 6:1-27\r\nc. Nelegiuirea lui Acan, 7:1-26\r\nd. Luarea cetatii Ai, 8:1-29\r\ne. Altarul de pe muntele Ebal, 8:30-35\r\n\r\nCAMPANIA SUDICa\r\na. Împaratii tarii se unesc 9:1 -2\r\nb. Viclenia Gabaonitilor, 9:3-27\r\nc. Marea lupta de la Gabaon, 10:1-11\r\nd. Soarele si luna se opresc, 10:12-15\r\ne. Uciderea celor cinci împarati, 10:16-27\r\nf. Alte izbînzi ale lui Iosua, 10:28-43\r\n\r\nCAMPANIA NORDICa\r\na. Marea batalie de la apele Meron, 11:1-15\r\nb. Cuceririle lui Iosua, 11:16-12:24\r\n\r\nIII. ÎMPaRtIREA taRII 13-24\r\na. Teritorii necucerite, 13:1-7\r\nb. Împartirea Transiordaniei, 13:8-33\r\n\r\nÎMPaRtIREA CANAANULUI\r\na. tinutul semintiei lui Iuda, 14:1-15:63\r\nb. tinutul semintiei lui Beniamin, 16:1-10\r\nc. tinutul semintiei lui Manase, 17:1-18\r\nd. Masurarea teritoriului ramas, 18:1-10\r\ne. tinutul lui Beniamin, 18:11-28\r\nf. tinutul lui Simeon, 19:1-9\r\ng. tinutul lui Zabulon, 19:10-16\r\nh. tinutul lui Isahar, 19:17-23\r\ni. tinutul lui Aser, 19:24-31\r\nî. tinutul lui Neftali, 19:32-39\r\nj. tinutul lui Dan, 19:40-48\r\nk Partea lui Iosua, 19:49-51\r\nl. Cetatile de scapare, 20:1-9\r\nm. Cetatile Levitilor, 21:1-45\r\nn. Semintiile din Transiordania, 22:1-34\r\no. Ultimele cuvinte ale lui Iosua, 23:1-16\r\np. Ultima adunare, 24:1-28\r\nr. Moartea lui Iosua, 24:29-33" },
    //JUDECĂTORI
    { "JUDECĂTORI", "Explicatia Cartii: Judecatori \r\n\r\nCu cartea Judecatori ajungem sa facem cunostinta cu unul dintre cele mai remarcabile personaje din istoria lui Israel: profetul-judecator Samuel. Acest om extraordinar a fost o figura pivotala în trecerea Israelului de la sistemul teocratie la cel monarhic. Cartile scrise de el fac legatura naturala între timpul lui Iosua si vremea împaratilor. Pentru a întelege cartile Judecatori si Rut trebuie sa plecam de la criza sfîsietoare prin care a trecut Samuel spre sfîrsitul vietii lui. În l Samuel 8:1-3 ni se spune ca fii lui Samuel n-au calcat pe urmele tatalui lor si poporul nemultumit de comportamentul lor ca judecatori s-a strîns la Rama si i-a cerut sa aseze peste natiune un împarat „cum au toate neamurile\" (1 Sam. 8:5). Cererea aceasta l-a durut foarte mult pe Samuel care s-a vazut lepadat de popor înainte de vreme. Dumnezeu i s-a aratat însa în mîhnirea lui si i-a explicat ca cererea poporului era mult mai vinovata decît parea la prima vedere: „Asculta glasul poporului în tot ce-ti va spune; caci nu pe tine te leapada, ci pe Mine Ma leapada, ca sa nu mai domnesc peste ei. Ei se poarta cu tine cum s-au purtat totdeauna, de cînd i-am scos din Egipt pîna în ziua de astazi; M-au parasit si au slujit altor dumnezei\", (1 Sam. 8:7-8) Desi fusese unul dintre cei mai buni judecatori din Israel, Samuel a trait drama ca sub conducerea lui poporul sa se lepede de Dumnezeu si sa ceara un om muritor drept împarat peste natiune. Samuel li l-a dat atunci pe Saul (1 Sam. 9 si 10). O buna bucata de timp, Samuel a fost torturat mereu de o întrebare chinuitoare: „Unde am gresit?\" Ca sa scape de ea, el i-a mai adunat o data la un loc pe toti evreii si le-a spus: „Iata-ma! Marturisiti împotriva mea... : Cui i-am luat boul, sau cui i-am luat magarul? Pe cine am asuprit si pe cine am napastuit? De la cine am luat mita ca sa închid ochii asupra lui? Marturisiti si voi da înapoi\". Ei au raspuns: „Nu ne-ai apasat, nu ne-ai napastuit, si nici n-ai primit nimic din mîna nimanui\". Atunci Samuel a zis poporului:...” (1 Sam. 12:3-17) Ceea ce urmeaza în text este o recapitulare selectiva a istoriei lui Israel cu scopul de a arata poporului tragedia alegerii pe care au facut-o. Dumnezeu însusi a confirmat din cer spusele lui Samuel dînd, în mijlocul anotimpului secetos, „chiar în ziua aceea tunete si ploaie\" (1 Sam. 12:17-18). Speriat de cele întîmplate poporul s-a cait si a vrut sa retracteze cererea pentru un împarat. Dar era prea tîrziu. Dumnezeu facuse pe placul poporului, îl aveau de acum pe Saul.\r\n\r\nCe se va întîmpla însa cu Samuel? Amarît în suflet, el va ramîne credincios acestui popor îndaratnic si va cauta sa-i învete în continuare prin intermediul scrisului. Discursul sau magistral, va fi dezvoltat si îmbogatit cu alte evenimente din trecut devenind o veritabila lectie de pedagogie istorica. Retras în singuratate, Samuel se asterne la lucru si scrie cartea Judecatori, cartea Rut si cartile 1 si 2 Samuel, caci spusese el: „Departe de mine sa pacatuiesc împotriva Domnului, încetînd sa ma rog pentru voi! Va voi învata calea cea buna si dreapta\", (1 Sam. 12:23)\r\nCartea îsi ia titlul din însusi cuprinsul ei care este dedicat perioadei asa numitilor „judecatori\", reprezentantii lui Iehova, împaratul nevazut al Israelului. Acesti judecatori (sofetim - în ebraica) nu reprezentau o succesiune ordonata de guvernatori, ci niste eliberatori ocazionali ridicati de Dumnezeu pentru izbavirea lui Israel si pentru administrarea dreptatii.\r\nAsa cum am vazut, materialul cartii este pus împreuna de Samuel. Asta nu înseamna ca toate fragmentele cronicii istorice alcatuite de el au fost initial opera lui. Probabil ca multe din partile acestei carti au fost cronici ale semintiilor lui Israel. Amanuntele prezente în texte ca fabula lui Iotam, cîntarea Deborei, mesajul lui Iefta catre regele Amon si detaliile marunte din descrierea adunarii de la Mitpa ne trimit la niste autori care au fost ei însisi martori participanti la aceste evenimente. În acelasi timp, apare evident faptul ca aceste relatari au fost editate mai tîrziu de cineva care le-a pus împreuna. Aflam astfel ca la scrierea cartii, chivotul legamîntului fusese luat deja din Silo (Jud. 18:31; 20:27) si ca Israelul trecuse deja de pragul instaurarii monarhiei (17:7; 18:1; 19:1; 21:25). Nu trecuse însa vremea lui David, caci despre Iebusiti ni se spune ca: „au locuit în Ierusalim cu fii lui Beniamin pîna în ziua de azi\" (Jud. 1:21)\r\nPe vremea lui Saul, dupa 930 î.Cr. Evenimentele cartii acopere însa primii 300 de ani petrecuti de Israel în Canaan, aproximativ între anii 1380 -1050 î.Cr.\r\nSamuel face ca evenimentele acestei perioade sa sune ca o lectie de istorie pentru poporul Israel. Dorinta lui este sa-i convinga pe evrei ca din orice greseala exista posibilitatea întoarcerii la „Dumnezeul care nu oboseste iertînd\". Desi sînt pomeniti un total de 12 judecatori, numai despre cîtiva se vorbeste în detaliu asa ca tot cuprinsul cartii poate fi rezumat la cinci serii de sapte: 7 apostazii, urmate de 7 pedepse prin robie, care fac poporul sa se pocaiasca de 7 ori si se termina prin 7 izbaviri minunate lucrate de Dumnezeu prin intermediul a 7 judecatori. Samuel stia si el ca repetitia este mama învataturii. De aceea el alege din trecutul Israelului 7 situatii în care poporul a gresit înaintea lui Dumnezeu si le arata cum, de fiecare data, pocainta a fost singura iesire posibila de sub pedeapsa divina. Urmarea acestei pocainte a fost, de fiecare data, îndurarea lui Dumnezeu, izbavirea si binecuvîntarea, în felul în care o reda Samuel, istoria lui Israel este o istorie ciclica.\r\n\r\nCuvinte cheie si teme caracteristice: Exista o fraza-comentariu care se repeta ca un advertisment în punctele de greutate ale cartii: „În vremea aceea nu era împarat în Israel. Fiecare facea ce-i placea\" (17:6; 18:1; 21:25). Exista aici un semnal de alarma pentru cei ce nu-L vor pe Dumnezeu sa domneasca asupra lor si nu accepta Cuvîntul Lui ca norma cu autoritate absoluta asupra vietii lor.\r\n\r\nFiecare din ratacirile lui Israel este introdusa cu aceleasi cuvinte: „Copiii lui Israel au facut ce nu placea Domnului...” (3:7; 3:12; 4:1; 6:1; 10:6; 13:1). Pedepsele trimise de Dumnezeu pentru neascultarea lor sînt numite cu termeni care subliniaza dreptul de posesiune al lui Iehova asupra Israelului: „ªi Domnul i-a vîndut în mîinile lui...” (3:8; 4:2; 10:7) sau „ªi Domnul i-a dat...” (6:1; 13:1). Pocainta poporului sub chinul pedepsei este descrisa mereu cu aceleasi cuvinte: „Copiii lui Israel au strigat catre Domnul\" (3:9; 3:15; 4:3; 6:7; 10:10). Repetarea aceasta obsedanta scoate în evidenta caracterul ciclic al oscilatiilor poporului Israel si mareste probabilitatea ca lectia data de Samuel sa fie înteleasa bine de cititorii lui evrei.\r\n\r\nPentru cititorul crestin, cartea Judecatori mai ilustreaza si alte teme biblice:\r\n\r\n1. Rautatea si depravarea inimii umane „deznadajduit de rea si de înselatoare\" (Jud. 2:11-13, 17, 19; 8:33-35; 10:6; 13:1).\r\n\r\n2. Placerea pe care o are Dumnezeu sa se foloseasca de lucruri slabe: Ehud - un infirm stîngaci (3:15), ªamgar - cu un otig de plug (Jud. 3:31), Debora - o femeie (Jud. 4:2, 9, 21; vezi si Jud. 9:53, Ghedeon - cel mai mic dintr-o familie saraca (Jud. 6:15), ceata mica de ostasi cu ulcioarele în mîini (Jud. 7:16) si falca de magar din mîna lui Samson (Jud. 15:15).\r\n\r\n3. Duhul Sfînt este cel ce da tarie si biruinta, în viata lui Otniel (Jud. 3:10), în viata lui Ghedeon (Jud. 6:34), în viata lui Iefta (Jud. 11:29), în viata lui Samson (Jud. 13:25; 14:6; 15:14), în viata tuturor copiilor lui Dumnezeu (Zaharia 4:6).\r\n\r\nCartea Judecatorilor se termina într-un dezastru total si generalizat. Stricaciunea morala, civica si spirituala prevala pretutindeni. Asta nu înseamna ca poporul nu avea totusi o lege. Falimentul oamenilor nu anuleaza standardul lui Dumnezeu, tot asa cum astazi dispretul fata de cele 10 porunci si fata de învataturile lui Cristos nu înseamna ca acestea nu exista. Cartea Judecatori este o dovada ca, lasat de capul lui, omul nu merge în sus pe calea progresului, ci coboara vertiginos în abisurile pacatului. Lumea are nevoie de izbavire. Ca si alta data, oriunde exista pocainta sincera ea se va grabi sa soseasca. ªi va aduce o data cu ea si binecuvîntarea!\r\n\r\nMesajul cartii: Este clar ca Samuel nu a fost preocupat sa redea fidel si cronologic istoria. Dorinta lui este sa noteze numai anumite evenimente care au o semnificatie spirituala aparte. Aceasta explica de ce unor întîmplari li se acorda un spatiu atît de mare în naratiune, iar altora doar un spatiu restrîns. Lipsesc aproape cu desavîrsire referintele la activitatea „marilor preoti\" care au fost activi în aceasta perioada, iar istoriile judecatorilor nu sînt asezate neaparat într-o ordine naturala. De fapt, unii dintre ei au functionat simultan în teritorii diferite din întinderea Israelului. Mesajul cartii trebuie cautat în ultimul discurs tinut de Samuel în fata poporului: „Nu va temeti! Ati facut tot raul acesta; dar nu va abateti de la Domnul, si slujiti Domnului din toata inima voastra. Nu va abateti de la El; altfel ati merge dupa lucruri de nimic, care n-aduc nici folos, nici izbavire, pentru ca sînt lucruri de nimic. Domnul nu va parasi pe poporul Lui, din pricina Numelui Lui cel mare, caci Domnul a hotarît sa faca din voi poporul Lui...  Temeti-va numai de Domnul, si slujiti-L cu credinciosie din toata inima voastra; caci vedeti ce „putere desfasoara El printre voi\" (1 Sam. 12:20-24)\r\n\r\nSCHItA CaRtII\r\n\r\nCADRUL ISTORIC GENERAL 1:1-3:6\r\na. Situatia politica, 1:1-36\r\nb. Situatia spirituala, 2:1-3:6\r\n\r\nI. PRIMUL CICLU 3:7-11\r\na. Pacatul idolatriei, 3:7\r\nb. Pedeapsa prin robie, 3:8\r\nc. Pocainta evreilor, 3:9\r\nd. Ridicarea lui Otniel, 3:9-10\r\ne. tara are odihna, 3: 1 1\r\n\r\nII. AL DOILEA CICLU 3:12-31\r\na. Neascultarea poporului, 3:12\r\nb. Pedeapsa prin robie, 3:12-14\r\nc. Pocainta evreilor, 3:15\r\nd. Ridicarea lui Ehud, 3:12-29\r\ne. tara are odihna, 3:30\r\n\r\nIII. AL TREILEA CICLU 4 - 5\r\na. Neascultarea poporului, 4:1\r\nb. Pedeapsa prin robie, 4:2\r\nc. Pocainta evreilor, 4:3\r\nd. Debora si Barac, 4:4-5:31\r\ne. tara are odihna, 5:31\r\n\r\nIV. AL PATRULEA CICLU 6:1 - 8:32\r\na. Neascultarea poporului, 6:1\r\nb. Pedeapsa prin robie, 6:1-5\r\nc. Pocainta evreilor, 6:6\r\nd. Mustrarea, 6:7-10\r\ne. Ridicarea lui Ghedeon, 6:11-24\r\nf. Ghedeon darîma altarul lui Baal, 6:25-32\r\ng. Ghedeon izbaveste poporul, 6:33-8:32\r\n\r\nV. AL CINCILEA CICLU 8:33 - 10:5\r\na. Pacatul idolatriei, 8:33-35\r\nb. Abimelec vrea sa fie împarat, 9:1-57\r\nc. Ridicarea altor judecatori, 10:1-5\r\n\r\nVI. AL ªASELEA CICLU 10:6 - 12:5\r\na. Pacatul idolatriei, 10:6\r\nb. Pedeapsa robiei, 10:7-9\r\nc. Pocainta evreilor, 10:10\r\nd. Mustrarea Domnului, 10:11-14\r\ne. Cererea poporului, 11:15\r\nf. Ridicarea lui Iefta, 11:16-12:7\r\ng. Alti judecatori, 12:8-15\r\n\r\nVII. AL ªAPTELEA CICLU 13-16\r\na. Neascultarea poporului, 13:1\r\nb. Pedeapsa robiei, 13:1\r\nc. Ridicarea lui Samson, 13:1-25\r\nd. Hartuielile lui Samson cu filistenii 14-15\r\ne. Samson si Dalila, 15:1-21\r\nf. Razbunarea si moartea lui Samson, 15:22-31\r\n\r\nPaCATUL SE EXTINDE, 17-21\r\na. Mica schimba preotia, 17:1-13\r\nb. Danitii preiau apostazia, 18:1-31\r\nc. Depravarea beniamitilor, 19:1-21:25\r\nCauta in biblie\r\nTextul cautat\r\nCauta...\r\nCum caut\r\nToate Cuvintele\r\nUnde caut:\r\nToată biblia\r\nVersetul Zilei\r\nDaniel 5:26-28\r\n\r\nIata insa scrierea care a fost scrisa: Mene - DUMNEZEU ti-a numarat zilele domniei, si i-a pus capat. Techel - inseamna ca ai fost cantarit in cumpana si ai fost gasit usor. Peres - inseamna ca imparatia ta va fi impartita, si data Mezilor si Persilor.\r\nMeditatia Zilei\r\n\r\nPowered by Biblia Online\r\nFacebook\r\n\r\nParteneri\r\nStiri Crestine\r\nMuzica Crestina\r\nVersuri Crestine\r\nChristian Lyrics\r\nBiblia Online Cornilescu\r\nVersuri Crestine\r\nNu este Crăciun fără Isus - Estera & Laura Bretan\r\nÎn iesle azi - Diana Scridon Pop\r\nBucurie in suflet - Familia Timofte\r\nSe-aude Glas Peste Carpați / Tăria Noastră Fii Mereu - The Beuca Family\r\nUn singur Dumnezeu - Otto Pascal & Biji\r\nFiu iubit - BBSO\r\nClipa de clipa - Philadelphia Band\r\nPrintul Pacii a venit - Alin și Emima Timofte & TB Music Family\r\nPraise - Lumina Worship\r\nUnde ești țara mea?\r\nPowered by VersuriCrestine.ro\r\n \r\n" },
    //RUT
    { "RUT", "Explicatia Cartii: Rut \r\n\r\nCartea „Rut\" este o oază de credincioşie într-o epocă caracterizată de idolatrie şi apostazie.\r\n\r\nTitlul: Este una dintre singurele două cărţi din Biblie care poartă numele unei femei. Cealaltă este „Estera\" cu care se contrastează şi se completează reciproc.\r\n\r\nAutorul: Samuel a scris această carte probabil după ce l-a uns pe David ca nou împărat peste Israel.\r\n\r\nData: Aproximativ în jurul anului 1.000 î.Cr.\r\n\r\nConţinutul cărţii: Cartea Rut este un apendice al cărţii Judecătorilor. De fapt, ea ar fi trebuit să fie inclusă în acea carte. Iată ce declară autorul ei în primul verset al cărţii: „Pe vremea judecătorilor a fost o foamete în ţară...”\r\n\r\nDe ce a scos Samuel această întîmplare din cartea Judecătorilor şi a făcut-o o carte de sine stătătoare? Motivele sînt două şi amîndouă sînt clare ca lumina zilei.\r\n\r\nÎn primul rînd, Samuel a vrut ca această carte să fie o încheiere recapitulativă a lecţiilor date în cartea Judecători, încă o dată ne vom întîlni aici cu ciclul de apostazie, pedeapsă, pocăinţă şi izbăvire. De data aceasta însă, Samuel vrea să răspundă unei întrebări viclene: „Şi dacă nu ne vom pocăi? Ce se va întîmpla cu noi dacă în loc să ne pocăim ne vom îndepărta şi mai mult de Domnul şi vom fugi pe alte meleaguri, departe de faţa Sa?\" Experienţa tristă a familiei lui Elimelec este răspunsul clar şi răspicat dat acestor întrebări răuvoitoare: „Nicăieri nu este prea departe pentru mîna Domnului ca să ne ajungă şi să ne pedepsească\". Nu există altă soluţie de ieşire de sub pedeapsă decît pocăinţa. Cînd ne pocăim cu adevărat Dumnezeu se grăbeşte să ne dea izbăvirea şi binecuvîntarea.\r\n\r\nÎn al doilea rînd, istoria familiei lui Elimelec este scoasă în evidenţă pentru că din această familie s-a născut mai tîrziu Işai, tatăl lui David, acest nou împărat uns de Samuel peste Israel. Providenţial, scrierea lui Samuel este mai importantă decît şi-a dat el atunci seama căci ea ne arată o secvenţă din viaţa familiei care-L va da lumii pe Mesia: Domnul Isus Cristos din seminţia lui David.\r\n\r\nCuvinte cheie şi teme caracteristice: O citire foarte atentă va scoate la iveală un mesaj spiritual latent ascuns în semnificaţia numelor purtate de personajele cărţii, întîmplarea debutează la Betleem, care tradus înseamnă „casa plinii\". Primul personaj cu care ne întîlnim este Elimelec al cărui nume se traduce prin „Dumnezeu este împăratul meu\" (Eli=Dumnezeu, melech=împărat). Acest bărbat pleacă din Betleem împreună cu nevasta lui numită Naomi, care se poate traduce prin „plăcuta\" sau „fermecătoarea\", şi cu cei doi fii ai lor: Mahlon, care înseamnă „cîntec\" şi „bucurie\", şi Chilion, care înseamnă „ornament\" sau „desăvârşire\". Hotărîrea lor de a pleca este echivalentă cu o încercare de a scăpa de sub pedeapsa pe care Dumnezeu o trimisese asupra Israelului prin secetă, în Moab, Elimelec („Dumnezeu este împăratul meu\") moare şi după el mor şi cei doi fii: Mahlon (bucuria) şi Chilion (desăvârşirea). Frîntă de durere, Naomi hotărăşte să se întoarcă în Israel, dar refuză să mai fie numită Naomi şi cere să i se spună de acuma: Mara, care înseamnă amărăciune. Mesajul spiritual din această succesiune de evenimente este foarte clar: în Canaan, Israelul trăia sub teocraţie („Dumnezeu este împăratul meu\") şi făcea casă bună cu plăcerea, cîntecul, bucuria, frumuseţea şi desăvîrşirea. Sub pedeapsa lui Dumnezeu, ei au recurs la soluţii lăturalnice, refuzînd pocăinţa şi asta le-a adus moarte, suferinţă şi sărăcie. Pentru copiii Domnului, singura ieşire din impasul neascultării este pocăinţa şi întoarcerea acasă, la Dumnezeului lor care nu oboseşte iertîndu-i şi reparîndu-le greşelile. Din momentul hotărîrii de întoarcere în ţară, mîna nevăzutului Dumnezeu începe să lucreze tainic pentru refacerea viitorului lui Naomi. Una dintre nurori, moabită Rut, se întoarce împreună cu ea şi găseşte favoare în ochii unui om bogat numit Boaz (tradus prin: „în El este puterea\"). Acest Boaz restaurează moştenirea lui Naomi şi-i dăruieşte un fiu prin Rut. Fiul lui Naomi din Rut, Obed, fusese hotărît de Dumnezeu să fie tatăl lui Işai şi bunicul lui David, împăratul slăvit al lui Israel.\r\n\r\nO altă temă foarte importantă din cartea Rut este„Goel\", ruda cu drept de răscumpărare. În rînduială lui Israel, cineva căzut în robie putea fi răscumpărat de un astfel de Goel care trebuia să îndeplinească 3 condiţii: să fie cea mai apropiată rudă, să vrea să răscumpere şi să poată să răscumpere, în cazul lui Naomi au existat două persoane care aveau acest drept: una al cărui nume nu este specificat (Rut 3:12; 4:1-6) şi Boaz (2:20; 4:9-13). Prin răscumpărarea pe care o face, Boaz este un tip care-l prevesteşte pe Domnul Isus. Lucrarea Lui de răscumpărare este descrisă strălucit în Apocalipsa 5:1-10:\r\n\r\n„Apoi am văzut în mîna dreaptă a Celui ce şedea pe scaunul de domnie o carte, scrisă pe dinlăuntru, şi pe dinafară, pecetluită cu şapte peceţi. Şi am văzut un înger puternic, care striga cu glas tare: „Cine este vrednic să deschidă cartea şi să-i rupă peceţile?\" Şi nu se găsea nimeni nici în cer, nici pe pămînt, nici sub pămînt, care să poată să deschidă cartea, nici să se uite în ea. Şi am plîns mult, pentru că nimeni nu fusese găsit vrednic să deschidă cartea şi să se uite în ea.\r\n\r\nŞi unul dintre bătrîni mi-a zis: „Nu plînge: Iată că Leul din seminţia lui Iuda, Rădăcina lui David, a biruit ca să deschidă cartea, şi cele şapte peceţi ale ei. Şi la mijloc, între scaunul de domnie şi cele patru făpturi vii, şi între bătrîni, am văzut stînd în picioare un Miel. Părea junghiat, şi avea şapte coarne şi şapte ochi, care sînt cele şapte Duhuri ale lui Dumnezeu, trimise în tot pămîntul. El a venit, şi a luat cartea din mîna dreaptă a Celui ce şedea pe scaunul de domnie. Cînd a luat cartea, cele patru făpturi vii şi cei douăzeci şi patru de bătrîni s-au aruncat la pămînt înaintea Mielului, avînd fiecare cîte o alăută şi potire de aur, pline cu tămîie, care sînt rugăciunile sfinţilor. Şi cîntau o cîntare nouă, şi ziceau: „Vrednic eşti tu să iei cartea şi să-i rupi peceţile, căci ai fost jungheat şi ai răscumpărat pentru Dumnezeu, cu sîngele Tău, oameni din orice seminţie, de orice limbă, din orice norod şi de orice neam. Ai făcut din ei o împărăţie şi preoţi pentru Dumnezeul nostru, şi ei vor împăraţi pe pămînt\".\r\n\r\nPentru că putea şi dorea să ne răscumpere, Domnul Isus s-a înrudit cu noi prin întrupare şi a devenit Goel-ul nostru.\r\n\r\nPe cine simbolizează însă cel care, deşi putea, n-a vrut să facă răscumpărarea Naomei din cauza moabitei Rut? Ei bine răspunsul este simplu. Acel „cutare\" (Rut 4:1) simbolizează pretenţiile Legii mozaice. Textul Legii spunea: „Amonitul şi Moabitul să nu intre în adunarea Domnului, nici chiar al zecilea neam, pe vecie. Să nu-ţi pese nici de propăşirea lor, nici de bună starea lor, toată viaţa ta, pe vecie\". (Deut. 23:3, 6) Moabită Rut este un simbol al Bisericii. Dumnezeu ne-a răscumpărat prin Cristos, dincolo de prevederile Legii, printr-un act de har şi de nemăsurată îndurare.\r\n\r\nSCHIŢA CĂRŢII\r\n\r\nI. ZADARNICA FUGA DE PEDEAPSĂ 1:1-5\r\na. Contextul istoric, 1:1\r\nb. Contextul familial, 1:2\r\nc. Rezultatul fugii, 1:3-5\r\n\r\nII. ÎNTOARCEREA 1:6-22\r\na. Hotărîrea Naomei, 1:6-7\r\nb. Sfatul dat nurorilor, 1:8-13\r\nc. Hotărîrea lui Rut, 1:14-19\r\nd. Naomi se cheamă acum Mara, 1:20-22\r\n\r\nIII. ÎNCEP COINCIDENŢELE 2:1-23\r\na. Rut merge pe ogorul lui Boaz, 2:1-3\r\nb. Rut este lăudată de Boaz, 2:4-13\r\nc. Rut capătă trecere înaintea lui Boaz, 2:14-16\r\nd. Naomi întrevede binecuvîntarea, 2:17-22\r\ne. Rut rămîne în ogoarele lui Boaz, 2:23\r\n\r\nIV. ÎNTÎLNIREA DIN NOAPTE, 3:1-18\r\na. Sfatul dat lui Rut de Naomi, 3:1-4\r\nb. Ascultarea lui Rut, 3:5-6\r\nc. Boaz înţelege şi acceptă, 3:7-15\r\nd. Naomi are credinţă, 3:16-18\r\n\r\nV. ÎNTÎLNIREA DE LA POARTĂ 4:1-12\r\na. Cererea lui Boaz, 4:1-4\r\nb. Explicaţia lui Boaz, 4:5\r\nc. Refuzul răscumpărătorului, 4:6\r\nd. Boaz oferă răscumpărare, 4:7-10\r\ne. Poporul se bucură, 4:11-12\r\n\r\nVI. DUMNEZEU DĂ UN VIITOR LUI NAOMI 4:13-22\r\na. Naşterea copilului, 4:13\r\nb. Urarea femeilor din cetate, 4:14-17\r\nc. Linia davidică, 4:18-21\r\nCauta in biblie\r\nTextul cautat\r\nCauta...\r\nCum caut\r\nToate Cuvintele\r\nUnde caut:\r\nToată biblia\r\nVersetul Zilei\r\nDaniel 5:26-28\r\n\r\nIata insa scrierea care a fost scrisa: Mene - DUMNEZEU ti-a numarat zilele domniei, si i-a pus capat. Techel - inseamna ca ai fost cantarit in cumpana si ai fost gasit usor. Peres - inseamna ca imparatia ta va fi impartita, si data Mezilor si Persilor.\r\nMeditatia Zilei\r\n\r\nPowered by Biblia Online\r\nFacebook\r\n\r\nParteneri\r\nStiri Crestine\r\nMuzica Crestina\r\nVersuri Crestine\r\nChristian Lyrics\r\nBiblia Online Cornilescu\r\nVersuri Crestine\r\nNu este Crăciun fără Isus - Estera & Laura Bretan\r\nÎn iesle azi - Diana Scridon Pop\r\nBucurie in suflet - Familia Timofte\r\nSe-aude Glas Peste Carpați / Tăria Noastră Fii Mereu - The Beuca Family\r\nUn singur Dumnezeu - Otto Pascal & Biji\r\nFiu iubit - BBSO\r\nClipa de clipa - Philadelphia Band\r\nPrintul Pacii a venit - Alin și Emima Timofte & TB Music Family\r\nPraise - Lumina Worship\r\nUnde ești țara mea?\r\nPowered by VersuriCrestine.ro\r\n" },
    //1 SAMUEL
    { "1 SAMUEL", "Explicatia Cartii: 1 Samuel \r\n\r\nCu cartea 1 Samuel începem sa citim cele trei grupe de carti duble din Vechiul Testament - 1 si 2 Samuel, 1 si 2 Împarati si 1 si 2 Cronici. Ele alcatuiesc împreuna o unitate distincta acoperind perioada istorica delimitata de ridicarea si prabusirea monarhiei în Israel, adica aproximativ 500 de ani scursi pe albia vremii între 1095-586 î.Cr. Cartea poarta numele autorului ei: Samuel, care talmacit înseamna „cerut de la Dumnezeu\". În originalul evreiesc, cele doua carti ale lui Samuel, ca de altfel si celelalte doua carti duble, au format la început o singura carte, împartirea prezenta se datoreaza traducatorilor în limbile greaca si latina. Desi criticata de unii, împartirea actuala îsi are meritul ei, deoarece 2 Samuel, ocupîndu-se prin excelenta de cei 40 de ani de domnie ai împaratului David, merita sa alcatuiasca un volum aparte.\r\nFara nici o îndoiala, Samuel este cel ce a început sa scrie aceasta minunata cronica de istorie. Probabil ca primele 24 de capitole sînt în întregime ale lui. Restul a fost adaugat de Natan si Gad asa cum gasim lamurit în 1 Cronici 29:29 si în 1 Samuel 10:25.\r\nAproximativ 930 î.Cr. si cam 100 de ani dupa aceea. Cartea acopere o perioada de 115 ani din istoria lui Israel.\r\n1 Samuel este o cronica a tranzitiei lui Israel de la teocratie la monarhie. Dumnezeu n-a vrut niciodata ca Israelul sa aiba un alt împarat în afara de El însusi. Dorinta Lui a fost ca sa ridice din mijlocul poporului conducatori iscusiti care sa se afle întotdeauna sub ordinele Lui. Legea si Marele Preot ar fi trebuit sa reglementeze viata poporului în ascultarea lor deplina de Domnul. Dar n-a fost asa. Israelul s-a departat de Dumnezeu, a iubit pacatul si s-a aplecat spre idolatrie. Lipsiti de limitarile Legii poporul ajunsese fara frîu. Solutia ar fi fost o întoarcere la Domnul si la „marturie\", dar poporul a ales o alta cale: ei i-au cerut lui Samuel sa aseze un împarat peste ei, asa „cum ou toate neamurile\" (1 Sam. 10:3). Cererea lor fusese anticipata de previziunile lui Dumnezeu. Aduceti-va aminte ce le spusese El prin Moise în Deuteronomul 17:14-20. Poate ca într-un fel, batrînii lui Israel au vazut în aceast pasaj nu un avertisment, ci o ultima portita de iesire din impas pentru viitor.\r\n\r\nContinutul cartii 1 Samuel va putea fi tinut foarte usor în minte daca o vom numi: „cartea celor trei conducatori\": Samuel (cap 1-7), Saul (cap. 8-15) si David (cap.16-31). Fireste, actiunile acestor trei oameni se suprapun în anumite texte, dar, în mare, împartirea aceasta a cartii este corecta.\r\n\r\nSamuel este una ditre cele mai luminoase figuri din întreaga istorie a lui Israel. Cu greu am putea gasi scris ceva rau despre el în paginile Scripturii. Ca importanta, Samuel încheia sirul judecatorilor, este primul din sirul profetilor, întemeiaza prima miscare de educatie religioasa din istoria Israelului, întemeiaza monarhia asezînd pe tron primul împarat în persoana lui Saul, pentru ca mai apoi sa-l unga ca împarat pe David, cei care avea sa devina cel mai mare împarat din întreaga istorie a natiunii. Au mai existat sporadic si alti oameni despre care Biblia ne spune ca s-au numit profeti (Gen. 20:7; Num. 11:25; Jud. 6:8; Deut. 18:18), dar sirul profetilor biblici începe cu acest Samuel (Fapte 3:24; 13:20). El este cel care a întemeiat oficiul profetic si a înfiintat scolile profetilor (1 Sam. 10:5, 6, 11-12; 19:20).\r\n\r\nMai presus de orice, Samuel a fost însa un om al rugaciunii: el a fost nascut în urma staruintei în rugaciune (1 Sam. l:9-28), a fost un copil obisnuit cu rugaciunea (1 Sam. 3:1-19), a învatat poporul sa biruiasca prin rugaciune (7:5-10), cînd poporul a cerut un împarat, el a alergat la rugaciune (1 Sam. 8:6) si a continuat sa se roage pentru popor chiar si dupa ce l-au lepadat ca si conducator (1 Sam. 12:19-23). Lipsa unei vieti de rugaciune a fost privita de Samuel drept un pacat (1 Sam. 12:23).\r\n\r\nIata ce a realizat Dumnezeu prin viata lui Samuel: (1) i-a eliberat din robia filisteana, (2) i-a pregatit pentru împaratie, (3) le-a dat un loc permanent pentru chivot si (4) a întarit si mai mult preotia.\r\n\r\nSaul, cel dintîi împarat al Israelului, a fost una dintre cele mai tragice figuri dintre personajele Vechiului Testament. El a început foarte promitator, însa a cunoscut apoi un declin teribil si a sfîrsit-o rusinos si lamentabil. Deasupra întregii lui vieti a plutit un aer de rîvna amestecat cu neascultare. Modestia i s-a transformat în mîndrie, adevarul în minciuna, rusinea în nerusinare si bunatatea fata de dusmani în ura si ucidere de prieteni. Pe piatra de pe mormîntul acestui prim împarat al Israelului s-ar potrivi o fraza pe care a rostit-o chiar el însusi, într-una din rarele clipe de luciditate care i-au întrerupt tulburarea sufleteasca de la sfîrsitul vietii: „Am lucrat ca un nebun!\" (1 Sam. 26:21; vezi si 1 Sam. 13:13)\r\n\r\nDavid, „lumina ochilor lui Dumnezeu\", a fost una dintre cele mai marete personalitati din toate timpurile. El a îmbogatit coplesitor de mult istoria sociala, militara, religioasa si culturala a lui Israel. În 1 Samuel ne întîlnim cu David ca baietel la oi, cîntaret din arfa, purtator de arme, capetenie în lupta, ginere al împaratului, scriitor neîntrecut de psalmi, împarat uns de Samuel si... fugar ratacitor si nevinovat. Acest fiu al lui Iese si nepot al lui Boaz, din moabita Rut, a vazut lumina zilei în Betleem ca cel mai mic din cei 8 fii ai tatalui sau. La vîrsta de 18 ani a fost uns pentru prima data de Samuel ca împarat ales de Dumnezeu peste Israel si avea sa devina apoi cel mai important dintre toti împaratii lui Israel, întemeietorul dinastiei din care s-a nascut, „la plinirea vremii\", Isus Cristos - Mesia.\r\n\r\nFaima lui de cîntaret iscusit a ajuns la curtea împaratului Saul, unde a fost chemat sa intre în anturajul familiei regale. Prietenia lui cu Ionatan, fiul lui Saul, este una dintre cele mai frumoase si mai gingase pagini din istoria relatiilor dintre oameni. Cînd a fost promovat în fruntea armatei, vitejia si popularitatea lui printre evrei a stîrnit invidia împaratului care s-a hotarît sa-l omoare. Saul a încercat de 4 ori acest lucru (1 Sam. 19:10, 15, 20, 21, 23, 24). Dumnezeu avea însa alte planuri cu David si i-a ocrotit viata. Zbuciumul care a caracterizat aceasta perioada si urmarile lui în maturizarea deosebita a lui David pot fi vazute în continutul psalmilor 59 si 37.\r\n\r\nÎncercarile din tinerete au fost scoala prin care Dumnezeu îl pregatea pentru ceasul domniei peste Israel, înainte de a stapîni peste oameni, David a fost învatat sa fie stapîn pe el însusi si sa astepte, plin de încredere, initiativele si hotarîrile Domnului.\r\n\r\nCa pribeag, departe de amenintarile împaratului Saul, David a trait în mijlocul altor dusmani care amenintau sa-i ia viata. Din temerile acelor clipe s-a nascut psalmul 56 în care gasim credinta lui în Dumnezeu si totala lui abandonare în hotarîrile Celui Atotputernic, ajunse la o totala maturitate. 1 Samuel se termina cu moartea rusinoasa a lui Saul si ne lasa sa anticipam cu înfrigurare citirea lui 2 Samuel în care viata lui David va cunoaste împlinirea omului „dupa inima lui Dumnezeu\" (1 Sam. 13:14).\r\nDumnezeu a chemat Israelul la o relatie unica si speciala cu El în virtutea careia El s-a proclamat împaratul lor nevazut. Din pricina nepriceperii lor, ei au renuntat la acest privilegiu si au cerut un împarat uman care sa le mearga în frunte. Solutia lor a însemnat un mare pas înapoi în destinul lor istoric. Suferintele ulterioare si tragedia lor nationala ne stau drept marturie. Din ceea ce scrie în cartea 1 Samuel putem învata ca neacceptarea cailor lui Dumnezeu, neatîrnarea de prezenta Lui nevazuta, refuzarea trairii prin credinta de dragul comoditatii solutionarilor omenesti vor atrage întotdeauna asupra noastra suferinte costisitoare si un sfîrsit rusinos.\r\n\r\nSchita CartII\r\n\r\nI. SAMUEL, ULTIMUL JUDECATOR l - 7\r\n\r\nANA, MAMA LUI SAMUEL\r\na. necazul ei, 1:1-8\r\nb. rugaciunea ei, 1:9-18\r\nc. copilul ei, 1:19-23\r\nd. jertfa ei, 1:24-28\r\ne. cîntarea ei, 2:1-W\r\n\r\nSLUJIREA MICULUI SAMUEL\r\na. Situatia din Silo, 2:11-36\r\nb. Domnul i se descopere, 3:1-21\r\n\r\nRAZBOIUL CU FILISTENII\r\na. Înfringerea Israelului, 4:1-10\r\nb. Pierderea chivotului, 4:11-17\r\nc. Moartea lui Eli, I-Cabod, 4:18-22\r\nd. Pedeapsa asupra filistenilor, 5:1-12\r\ne. Filistenii trimit chivotul înapoi, 6:1-7:2\r\nf. Campania de refacere spirituala, 7:3-17\r\n\r\nII. SAUL, CEL DINTÎI ÎMPARAT 8-15\r\n\r\nRIDICAREA LUI SAUL\r\na. Israelitii cer un împarat, 8:1-22\r\nb. Alegerea lui Saul, 9:1-27\r\nc. Încoronarea lui Saul, 10:1-27\r\nd. Înfrîngerea Amonitilor, 11:1-15\r\ne. Samuel înfrunta natiunea, 12:1-25\r\n\r\nPRABUSIREA LUI SAUL\r\na. Razboiul cu filistenii, 13:1-7\r\nb. Jertfa nelegiuita, 13:8-9\r\nc. Sentinta rostita de Samuel, 13:10-14\r\nd. Situatia se înrautateste, 13:15-23\r\ne. Juramintele pripite, 14:1-52\r\nf. Ascultarile partiale, 15:1-35\r\n\r\nIII. DAVID, ALESUL DOMNULUI 16-31\r\n\r\nASCENSIUNEA LUI DAVID\r\na. David uns împarat, 16:1-13\r\nb. David adus la curtea lui Saul, 16:14-23\r\nc. David si Goliat, 17:1-58\r\nd. David si Ionatan, 18:1-4\r\ne. David pizmuit de Saul, 18:5-16\r\nf. David, ginerele împaratului, 18:17-30\r\n\r\nDAVID ESTE PRIGONIT DE SAUL\r\na. David aparat de Ionatan, 19:1-7\r\nb. David aparat de Mical, 19:8-17\r\nc. David aparat de Samuel, 19:18-24\r\nd. David, aparat de Ionatan, 20:1-42\r\ne. David aparat de Ahimelec, 21:1-9\r\nf. David aparat de Achis, 21:10-15\r\n\r\nDAVID Sl CEATA LUI\r\na. În pestera Adulam si la Mitpe, 22:1 -5\r\nb. Saul ucide preotii din Nob, 22:6-23\r\nc. David scapa Cheila, 23:1-12\r\nd. În pustia Zif si Maon, 23:15-29\r\ne. En-Ghendi. David cruta pe Saul, 24:1-22\r\nf. David si Nabal, 25:1-35\r\ng. David si Abigail, 25:36-44\r\nh. Pustia Zif. David cruta pe Saul, 26:1-25\r\n\r\nDAVID ÎN TARA FILISTENILOR\r\na. ªederea lui la tiglad, 27:1-12\r\nb. Filistenii se lupta cu Israelul, 28:1-4\r\nc. Saul cheama mortii, 28:5-14\r\nd. Mesajul lui Samuel, 28:15-19\r\ne. Groaza lui Saul, 28:20-25\r\nf. David este trimis înapoi, 29:1-11\r\ng. David îi biruieste pe Amaleciti, 30:1-19\r\nh. Împartirea prazii, 30:20-25\r\ni. Darurile trimise în Iuda, 30:26-31\r\n\r\nMOARTEA LUI SAUL\r\na. Israelul batut de filisteni, 31:1\r\nb. Moartea lui Saul si Ionatan, 31:2-10\r\nCauta in biblie\r\nTextul cautat\r\nCauta...\r\nCum caut\r\nToate Cuvintele\r\nUnde caut:\r\nToată biblia\r\nVersetul Zilei\r\nDaniel 5:26-28\r\n\r\nIata insa scrierea care a fost scrisa: Mene - DUMNEZEU ti-a numarat zilele domniei, si i-a pus capat. Techel - inseamna ca ai fost cantarit in cumpana si ai fost gasit usor. Peres - inseamna ca imparatia ta va fi impartita, si data Mezilor si Persilor.\r\nMeditatia Zilei\r\n\r\nPowered by Biblia Online\r\nFacebook\r\n\r\nParteneri\r\nStiri Crestine\r\nMuzica Crestina\r\nVersuri Crestine\r\nChristian Lyrics\r\nBiblia Online Cornilescu\r\nVersuri Crestine\r\nNu este Crăciun fără Isus - Estera & Laura Bretan\r\nÎn iesle azi - Diana Scridon Pop\r\nBucurie in suflet - Familia Timofte\r\nSe-aude Glas Peste Carpați / Tăria Noastră Fii Mereu - The Beuca Family\r\nUn singur Dumnezeu - Otto Pascal & Biji\r\nFiu iubit - BBSO\r\nClipa de clipa - Philadelphia Band\r\nPrintul Pacii a venit - Alin și Emima Timofte & TB Music Family\r\nPraise - Lumina Worship\r\nUnde ești țara mea?\r\nPowered by VersuriCrestine.ro\r\n" },
    //2 SAMUEL
    { "2 SAMUEL", "Explicatia Cartii: 2 Samuel\r\n\r\n1 Samuel ne-a prezentat falimentul împaratului ales de oameni: Saul. 2 Samuel ne aseaza înaintea ochilor frumoasa domnie a unui împarat „dupa inima lui Dumnezeu\": David. Singurul motiv pentru care cartea poarta numele de „2 Samuel\" este acela ca în original continutul ei a alcatuit o singura carte împreuna cu cartea „1 Samuel\".\r\nCronicarii care au pastrat scrisa amintirea acestor evenimente nu sînt cunoscuti cu certitudine, 1 Cronici 29:29-30 ni-i indica drept autori pe doi profeti: Natan si Gad.\r\nEste nesigura, plasata însa cu certitudine undeva în cuprinsul secolului 10 dinaintea lui Cristos.\r\n2 Samuel poate fi o carte usor de memorat daca o vom considera drept o carte biografica. Subiectul ei poate fi formulat astfel: 40 de ani de domnie din viata împaratului David (2 Sam. 5:4-5).\r\n\r\nDavid a preluat conducerea lui Israel într-un timp în care natiunea se afla într-o stare haotica. Moartea lui Saul fusese în acelasi timp si tragica si rusinoasa. Prima experienta a Israelului cu monarhia sfîrsise într-un total esec. Începutul cartii „2 Samuel\" ni-l arata pe David întors la ticlag dupa biruinta pe care o avusese asupra Amalecitilor (2 Sam.1:1). El era obosit în trup, dar bucuros în suflet pentru victoria repurtata. Grija lui cea mare însa era la soarta celeilalte batalii: aceea dintre filisteni si Israel. Dupa trei zile, David primeste vestea ca Israelul a fost învins. Durerea lui David este sincera si nemîngîiata. Cu aceasta ocazie iese în relief toata frumusetea caracterului sau. Cîntarea de jale scrisa de el pentru Saul si mai ales pentru Ionatan (1: 17-27) este o elegie a unui suflet îndragostit si neatins de morbul dusmaniei, în ciuda nebuniei ucigase a lui Saul, David l-a iubit pîna la capat.\r\n\r\nÎnceputurile domniei lui David au fost modeste si pline de momente de descurajare, dar împaratul a avut o credinta nestramutata în Dumnezeu. Caracteristica vietii lui David a fost o totala subordonare fata de initiativele si orarul lui Dumnezeu.\r\n\r\nSub conducerea lui David, Israelul a ajuns la apogeul puterii si stralucirii. Timpul domniei lui si timpul domniei fiului sau Solomon sînt cunoscute drept „epoca de aur a Israelului\". David a eradicat idolatria din tara. Conducerea lui militara a fost geniala si teritoriile cucerite de el au dus Israelul la cea mai mare întindere geografica din toata istoria lui. Faima si bogatia adusa de el evreilor era cunoscuta pretutindeni. Negustorii evrei cutreierau pamîntul între Nil, Tigru si Eufrat sau alunecau pe corabiile marilor pîna spre tarmuri exotice si îndepartate. Dumnezeu a rînduit ca pe vremea domniei lui David si Egiptul si puterea Mesopotamiei, cele doua puncte traditionale de dominatie mondiala, sa fie intrate într-o totala eclipsa. Pentru cîteva decenii, Israelul s-a aflat în fruntea natiunilor lumii. Puterea, cultura, faima si luxul aveau atunci o singura capitala: Ierusalimul. Tronul asezat în acest oras îsi exercita influenta asupra populatiei întregii lumii. De ce a facut Dumnezeu aceste lucruri pentru David? Raspunsul nu este greu de gasit. David a acceptat sa fie toata viata lui autoritatea numarul 2 în Israel. Secretul reusitelor lui este întoarcerea la spiritul teocratiei. Împaratul David s-a subordonat împaratului cerurilor cautînd toata viata lui sa faca numai „voia Domnului\" (Fapte 13:22; 1 Sam. 13:13-14). În inima lui David, Iehova a fost tot timpul adevaratul împarat al lui Israel. Pe sine s-a considerat doar un fel de reprezentant vizibil al divinitatii. Aduceti-va aminte de felul în care s-a purtat el cu Saul: de refuzul sau repetat de a-si face singur dreptate si respectul nelimitat fata de „unsul Domnului\" (1 Sam. 24:6; 26:10-11). Remarcati totala lui dependenta de indicatiile divine: s-a nagat atunci cînd a fost în momente de criza (1 Sam. 23:2, 10-12), s-a rugat atunci cînd a început domnia (2 Sam 2:1) si a asteptat hotarîrile lui Dumnezeu atunci cînd a fost prigonit si blestemat pe nedrept (2 Sam. 16:5-12). Probabil ca tocmai aceasta caracteristica a comportamentului lui a fost aceea care l-a facut pe Dumnezeu sa-l numeasca „un om dupa inima Lui\" (1 Sam. 13:13-14). Altfel este greu sa accepti aceasta numire pentru un om care a avut pacate si defecte care i-au întunecat ultimii ani ai vietii. Pentru cititorul atent, viata lui David se împarte în doua jumatati distincte. Cea de pîna la pacatuirea cu Bat-ªeba si cea de dupa aceasta cadere. Chiar si cartea le acorda fiecareia parti egale: 12 capitole pentru ascensiunea glorioasa a lui David si tot 12 capitole pentru necazurile care i-au marcat cea de a doua perioada a vietii.\r\n\r\nCuvinte cheie si teme caracteristice: În 2 Samuel 7:11-16 ne întîlnim cu legamîntul Davidic. Importanta lui se extinde nu numai asupra istoriei familiei lui David, ci si asupra istoriei întregii omeniri, mai ales în partea ei viitoare. Legamîntul Davidic este cheia divina pentru întelegerea planului lui Dumnezeu în istorie. Din momentul încheierii lui, evreii au stiut ca Mesia va veni din linia genealogica a lui David. Ei pastreaza aceiasi convingere si astazi (Isaia 11:1; Ieremia 23:5; Ezechiel 37:25; ca si Luca 1:30-33). Legamîntul în sine contine cîteva elemente distincte: (1) confirmarea tronului lui Israel. Pîna la acest pasaj, monarhia fusese o dorinta vinovata a evreilor. De acum înainte, tronul lui Israel intra în istorie ca o institutie divina, (2) consacrarea familiei lui David ca familie împarateasca. Lui David i se promit trei lucruri: o „casa\" cu semnificatia unei linii continue de descendenti, un „tron\" cu semnificatia autoritatii împaratesti si o „împaratie\" cu semnificatia unui domeniu de exercitare a autoritatii. 2 Samuel 7:16 ne declara ca promisiunea acestor trei elemente este „pe vecie\". Cel de al treilea element este: (3) anuntarea programului mesianic. Caracterul vesnic al celor trei promisiuni facute lui David implica împlinirea lor finala în viata lui Mesia (Isaia 9:7).\r\n\r\nLegamîntul Davidic marcheaza cea de a patra treapta de dezvoltare a promisiunilor mesianice. Lui Adam i s-a facut o primisiune care privea rasa umana în general; lui Avraam i s-a facut o promisiune care privea Israelul ca natiune; prin Iacov s-a facut o promisiune speciala semintiei lui Iuda, iar acum se identifica o familie din semintia lui Iuda: casa lui David. Ramîne doar ca Isaia sa mai adauge ca „samînta femeii\", fiul lui Avraam, Leul din semintia lui Iuda si odrasla lui David va fi nascut... dintr-o fecioara (Isaia 7:14).\r\nCele doua jumatati ale cartii ne spun clar ca exista o plata care trebuieste platita pentru pacat. Gresala lui David în cazul lui Bat-ªeba si Urie a atras asupra împaratului si asupra casei lui agonia pedepsei. O clipa de ratacire a dus la ani de suferinte. Dreptatea lui Dumnezeu se aplica fara discriminare asupra tuturor celor ce pacatuiesc, oricare ar fi pozitia lor înaintea Domnului. Exista însa chiar si în pedepsele Domnului o portie de har (2 Sam. 24:14).\r\n\r\nSCHItA CaRtII\r\n\r\nI. RIDICAREA LUI DAVID 1-12\r\n\r\nÎMPaRAT PESTE IUDA\r\na. Vestea morti lui Saul, 1:1-10\r\nb. Reactia lui David, 1:11-16\r\nc. Cîntarea de jale, 1:17-27\r\nd. David este uns împaratia Hebron, 2:1-7\r\ne. Is-Boset, rivalul lui David, 2:8-11\r\n\r\nRaZBOIUL CIVIL\r\na. Abner se lupta cu Ioab, 2:12-32\r\nb. Abner îl paraseste pe Is-Boset, 3:1-21\r\nc. Ioab îl omoara pe Abner, 3:22-39\r\nd. Is-Boset este asasinat, 4:1-12\r\n\r\nÎMPaRAT PESTE TOT ISRAELUL\r\na. David este uns împarat peste Israel, 5:1-6\r\nb. Stabilirea capitalei la Ierusalim, 5:7-25\r\nc. Chivotul este adus la Ierusalim, 6:1-23\r\n\r\nLEGaMÎNTUL DAVIDIC\r\na. Dorinta lui David, 7:1-3\r\nb. Promisiunea dumnezeiasca, 7:4-17\r\nc. Reactia lui David, 7:18-29\r\n\r\nCONSOLIDAREA ªl EXTINDEREA ÎMPaRatIEI\r\na. Biruintele lui David, 8:1-14\r\nb. Înaltii slujbasi ai lui David, 8:15-18\r\nc. David si Mefiboset, 9:1-13\r\nd. Biruinta asupra Amonitilor, 10:1-19\r\n\r\nPaCATUL LUI DAVID CU BAT-ªEBA\r\na. Comoditatea, 11:1\r\nb. Ispitirea, 11:2\r\nc. Pacatuirea, 11:3-5\r\nd. Perversitatea, 11:6-13\r\ne. Crima, 12:14-25\r\nf. Rezolvarea pacatoasa, 12:26-27\r\n\r\nII. NECAZURILE LUI DAVID 12-24\r\n\r\nDUMNEZEU MUSTRa ªl PEDEPSEªTE\r\na. Natan spune o întîmplare, 12:1-5\r\nb. Reactia lui David, 12:6\r\nc. Natan condamna pe fata, 12:7-9\r\nd. Rostirea sentintei, 12:10-12\r\ne. David se recunoaste vinovat, 12:13-14\r\nf. Moartea copilului, 12:15-23\r\ng. Nasterea lui Solomon, 12:24-25\r\nh. Refacerea autoritatii, 12:26-31\r\n\r\nNECAZURI ÎN FAMILIA LUI DAVID\r\na. Amnon curveste cu sora sa, 13:1-22\r\nb. Amnon este omorît de Absalom, 13:23-29\r\nc. Absalom fuge în exil, 13:34-39\r\nd. Ioab mijloceste pentru Absalom, 14:1-20\r\ne. Absalom este iertat si adus înapoi, 14:21 -33\r\nf. Rascoala lui Absalom, 15:1-12\r\ng. Fuga lui David, 15:13-37\r\nh. tiba înnegreste pe Mefiboset, 16:1-4\r\ni. ªimei bleastama pe David, 16:5-14\r\nî. Husai la Absalom, 16:15-19\r\nj. Absalom asculta de Ahitofel, 16:20-23\r\nk. Husai înfrunta pe Ahitofel, 17:1-14\r\nl. Moartea lui Ahitofel, 17:15-23\r\nm. David la Mahanaim, 17:24-18:4\r\nn. Absalom ucis de Ioab, 18:5-18\r\no. Jalea lui David, 18:19-33\r\n\r\nNECAZURI ÎN ÎMPaRatIE\r\na. Descumpanirea poporului, 19:1-8\r\nb. Refacerea armoniei în împaratie, 19:9-43\r\nc. Rascoala lui ªeba, 20:1 -3\r\nd. Amasa asasinat de Ioab, 20:4-13\r\nd. Uciderea lui ªeba, 20:14-22\r\ne. Dregatorii lui David, 20:23-26\r\nf. Foamete de trei ani, 21:1-14\r\ng. Filistenii pornesc alt razboi, 21:15-22\r\nh. Cîntarea lui David, 22:1-51\r\ni. David lauda legamîntul, 23:1-7\r\nî. Vitejii lui David, 23:8-39\r\nj. Numaratoarea poporului, 24:1-9\r\nk. Pedeapsa prin ciuma, 24:10-25" },
    //1 ÎMPARAȚI
    { "1 ÎMPĂRAȚI", "Explicatia Cartii: 1 Imparati\r\n\r\nCartile 1 si 2 Împarati cuprind o perioada de decadere si declin în istoria lui Israel. Ele încep cu împaratul David si termina cu împaratul Babilonului. Debuteaza cu constructia Templului si se încheie cu distrugerea lui. Încep cu o perioada de glorie si se termina în dezastru spiritual si robie nationala. Vremea împaratilor aduce în atentia noastra lucrarea profetilor biblici. l Împarati se termina cu lucrarea lui Ilie, acestui om de exceptie care a fost trimis de Dumnezeu sa întoarca poporul Israel din idolatrie la credinta cea adevarata. Stilul cartii arata lucrarea unui singur autor, chiar daca el a folosit si informatii din alte cronici ale vremii (1 Împarati 11:41; 14:29). Traditia evreiasca atribuie cartile acestea profetului Ieremia, în orice caz, cartile acestea au fost scrise înainte de distrugerea celui dintîi Templu.\r\nData: Aproximativ 550 î.Cr. Cartea acopere o perioada de 400 de ani de istorie înregistrînd cresterea si descresterea împaratiei lui Israel.\r\nÎn originalul ebraic, 1 si 2 Împarati au format o singura carte. Traducatorii în limba greaca le-au despartit în doua pentru ca textul grec era mai lung cu o treime si sulurile de papirus pe care scriau ei erau limitate în lungime. Semnificatia titlului este evidenta: aceste carti înregistreaza evenimentele din timpul domniei împaratului Solomon si ai celorlalti împarati care s-au succedat la tronul lui Israel si Iuda.\r\nÎn afara de timpul domniei lui Solomon, cartea 1 Împarati arata cum imediat dupa domnia lui Solomon, împaratia s-a divizat în alte doua împaratii: una în Nord continuînd sa poarte numele de Israel si una în sud cu numele de Iuda. Împaratia lui Israel cuprindea zece semintii si avea capitala la Samaria, în timp ce din împaratia lui Iuda au facut parte numai Iuda si Beniamin, iar capitala s-a aflat la Ierusalim. Pe tronul Samariei s-au perindat 19 împarati din semintii diferite. Ierusalimul a ramas credincios semintiei lui David din care s-au ridicat succesiv un numar de 20 de împarati.\r\n\r\nCa împartire, 1 Împarati se prezinta de la sine. Cele 22 de capitole sînt asezate în doua sectiuni egale. Primele 11 capitole sînt consacrate celor 40 de ani de domnie ai lui Solomon, iar celelalte 11 acopere primii 80 de ani de existenta separata a celor doua împaratii divizate.\r\n\r\nCuvinte cheia si teme caracteristice: Figura dominanta a cartii 1 Împarati este împaratul Solomon. Personalitatea lui este iesita din comun din trei puncte de vedere: ca înfaptuitor de istorie, ca influenta asupra societatii si ca semnificatie profetica mesianica.\r\n\r\nDin punct de vedere istoric, Solomon reprezinta momentul de apogeu al monarhiei evreiesti. Domnia lui marcheaza culmea de abundenta materiala si de cultura la care s-a ridicat vreodata Israelul. Cine citeste capitolele 9 si 10 ale cartii îsi da repede seama ca nivelul ajuns de Solomon a stîrnit uimirea lumii de atunci. Solomon a fost si ultimul împarat care a domnit peste întreg Israelul. Singurul care va mai face acest lucru va fi Isus Cristos, cînd se va întoarce sa domneasca ca fiu al lui David pe tronul Ierusalimului.\r\n\r\nCa lider social, Solomon este una dintre cele mai coplesitoare prezente din cîte au existat vreodata, întelepciunea lui super-normala l-a facut o celebritate admirata de toti contemporanii. A alcatuit trei mii de proverbe pline de discernamînt si de pricepere si a compus o mie cinci cîntari care-l aseaza în fruntea celor mai prolifici poeti din toate timpurile. Prietenii si dusmanii i-au cazut în admiratie si s-au grabit sa faca pace cu împaratia lui. Abilitatea sa organizatorica si administrativa a produs un salt de cîteva secole peste nivelul cunoscut în lumea si civilizatia de atunci. Capacitatea lui afectiva si emotionala a fost iesita din comun (vezi Cîntarea Cîntarilor). Excesele lui maritale nu pot fi explicate decît daca acceptam acest punct de vedere. Atitudinea lui fata de Dumnezeu este punctul fragil din întreaga lui constitutie. Bogatia nemasurata si neîmpartita cu saracii poporului, multimea nevestelor si numarul mare de care de lupta sînt toate în contrast evident cu prevederile Legii date de Dumnezeu evreilor (Deut. 17:16-17). Desi este elocvent si înfocat în rugaciune, se observa la el o oarecare detasare de Dumnezeu. Abundenta l-a facut ca încet, încet sa se socoteasca daca nu suficient în sine însusi, cel putin mai putin dependent de ajutorul divin, în comparatie cu modestia si sentimentul de neputinta pe care l-am remarcat la tatal sau David, Solomon prezinta o siguranta de sine si o opulenta care-i îneaca spiritul. Gîndurile sale exprimate în cartea Eclesiastul ni-l prezinta drept un cunoscator plictisit al tuturor dimensiunilor vietii umane. Blazarea lui este evidenta si suparatoare.\r\n\r\nFalimentul lui Solomon ca lider spiritual al poporului a demonstrat înca o data incapacitatea omului cazut în pacat de a sta într-o pozitie de totala autoritate. Nimeni n-a fost vreodata mai întelept ca Solomon si nimeni n-a avut la dispozitia sa mai multe resurse materiale si mai multa faima. Falimentul lui este reprezentativ si suficient pentru a explica falimentele noastre.\r\n\r\nApusul domniei lui Solomon prevesteste furtuna care se apropie. Tîrîta de iubire, inima lui îngaduie sa se cladeasca în Ierusalim temple pentru toti dumnezeii nevestelor lui. Chemos, Baal-Peor si Moloh intra cu fast în cetate si în inimile poporului. Israelul se prabuseste iarasi în idolatrie. Dumnezeu îl trimite pe proorocul Ahia sa-i vesteasca lui Solomon pedeapsa. Nimeni nu va mai aduna dupa el tot poporul într-o singura împaratie. Zece semintii vor trece sub stapînirea unui slujitor de al sau. Legamîntul Davidic însa va ramîne în picioare. David va avea mereu un urmas pe scaunul sau de domnie. Stapînirea lui va fi însa limitata la semintiile lui Iuda si Beniamin.\r\n\r\nSolomon este mult mai interesant si mai important însa ca semnificatie tipologica mesianica. Asemeni tatalui sau David, el este unul dintre cele mai clare anticipari ale lucrarii lui Isus Cristos. ªi tot ca David, Solomon este un simbol pentru aspectul viitor al domniei lui Cristos asupra lumii. Exista comentatori care vad în David aspectul domniei mesianice din timpul mileniului, iar în Solomon aspectul domniei din Noul Ierusalim de dupa mileniu. Asupra acestor detalii nu putem sa zabovim si nici sa ne pronuntam în acest scurt studiu.\r\n\r\nPrin ce simbolizeaza domnia lui Solomon aspecte din împaratia mesianica? Iata prin ce: Mai îmtîi, domnia lui a fost un timp de pace si odihna. Solomon n-a scos sabia din teaca în anii sai de domnie. În al doilea rînd, a fost o domnie caracterizata prin primatul întelepciunii si priceperii (1 Împarati 4 si 10). În al treilea rînd a fost vorba despre bogatie si slava - asa cum nu mai fusese nicaieri si niciodata, în al patrulea rînd a fost o domnie care s-a bucurat de faima si cinstire. Numele lui Solomon era renumit în toate împaratiile contemporane. În al cincilea rînd putem vorbi despre bucurie si siguranta. Iata ce putem citi în 1 Împarati 4:20,25: „Iuda si Israel erau în numar foarte mare... Ei mîncau si se veseleau...  au locuit în liniste fiecare sub via lui si sub smochinul lui, în tot timpul lui Solomon\".\r\n\r\nDin vestirile profetilor noi stim ca acestea sînt tocmai caracteristicile domniei mesianice viitoare a lui Cristos asupra lumii. Va fi pace si liniste: „Nici un popor nu va mai scoate sabia împotriva altuia, si nu vor mai învata razboiul\" (Isaia 2:4). „Atunci lupul va locui împreuna cu mielul, si pardosul se va culca împreuna cu iedul; vitelul, puiul de leu si vitele îngrasate, vor fi împreuna, si le va mîna un copilas\". (Isaia 11:6). Va fi un timp de întelepciune si pricepere asa cum n-a fost niciodata: „Pamîntul va fi plin de cunostinta Domnului ca fundul marii de apele care-l acopar\" (Isaia 11:9; Habacuc 2:14). Va fi bogatie si slava asa cum n-a mai fost niciodata caci: „muntele (împaratia) Casei Domnului va fi întemeiat ca cel mai înalt munte; se va înalta deasupra dealurilor, si toate neamurile se vor gramadi spre el\" (Isaia 2:2; 11:10J. Va fi slava si cinstire asa cum nici o împaratie n-a avut vreodata pe fata pamîntului: „În ziua aceea, Vlastarul lui Isai va fi ca un steag pentru popoare; neamurile se vor întoarce la El, si slava va fi locuinta Lui\" (Isaia 11:10). În împaratia mesianica va fi bucurie si siguranta pentru toti cei ce vor locui în ea. Profetul Mica scrie: „fiecare va locui sub vita lui si sub smochinul lui, si nimeni nu-l va mai tulbura. Caci gura Domnului ostirilor a vorbit\" (Mica 4:4). Nu exista nimic mai minunat decît studierea acestor pasaje care ne vorbesc despre Împaratia Domnului Isus care sta gata sa vina. Chiar si numai o lectura fugara a lor ne face sa spunem plini de dorinta: „Vie împaratia Ta!\"\r\n\r\nMesajul peste veacuri: Nu este greu de loc sa întelegem ce vrea Dumnezeu sa ne învete din cele întîmplate în cartea 1 Împarati. Textul care gazduieste acest mesaj este 1 Împ. 11:11-l3:\r\n\r\n„Fiindca ai facut asa, si n-ai pazit legamîntul Meu si legile Mele pe care ti le-am dat, voi rupe împaratia de la tine si o voi da slujitorului tau. Numai nu voi face lucrul acesta în timpul vietii tale, pentru tatal tau David. Ci din mîna fiului tau o voi rupe. Nu voi rupe însa toata împaratia; voi lasa o semintie fiului tau, din pricina robului Meu David, si din pricina Ierusalimului pe care l-am ales\".\r\n\r\nUnde nu exista ascultare, nu exista viitor. Pentru cel semet, Dumnezeu nu pastreaza nici un loc sub soare. Smeritul însa va ramîne pururi în harul dumnezeiesc si Dumnezeu îi va asigura vesnicia.\r\n\r\nSCHItA CaRtII\r\n\r\nI. DOMNIA LUI SOLOMON 1-11\r\na. Ridicarea lui Solomon, 1:1-3:1\r\n\r\nÎNtELEPCIUNEA LUI SOLOMON, 3:2-4:34\r\nSolomon cere întelepciune, 3:2-15\r\nSolomon dovedeste întelepciune, 3:16-28\r\nSolomon aseaza administratori, 4:1-28\r\nSolomon în marirea sa, 4:29-34\r\n\r\nb. Templul lui Solomon, 5:1-8:66\r\n\r\nFAIMA LUI SOLOMON, 9:1-10:29\r\nLegamântul cu Domnul, 9:1-9\r\nDarul dat lui Hiram, 9:10-14\r\nSlujitorii lui Solomon, 9:15-25\r\nFlota lui Solomon, 9:26-28\r\nVizita reginei din Seba, 10:1-13\r\nBogatiile lui Solomon, 10:14-29\r\n\r\nCaDEREA LUI SOLOMON, 11:1-43\r\na. Motivul, 11:1-8\r\nb. Avertizarea, 11:9-13\r\nc. Dusmanii, 11:14-28\r\nd. Profetia lui Ahia, 11:29-40\r\ne. Moartea lui Solomon, 11:41-43\r\n\r\nII. ÎMPaRatIILE DESPaRtITE 12-22\r\n\r\nDIVIZAREA ÎMPaRatIEI, 12:1-24\r\nCererea triburilor din nord, 12:1-4\r\nRaspunsul lui Roboam, 12:5-15\r\nRevolta triburilor din nord, 12:16-24\r\n\r\na. Domnia lui Ieroboam în Israel, 12:25-14:20\r\nb. Domnia lui Roboam în Iuda, 14:21-31\r\nc. Domnia lui Abiam în Iuda, 15:1-8\r\nd. Domnia lui Asa în Iuda, 15:9-24\r\ne. Domnia lui Nadab în Israel, 15:25-31\r\nf. Domnia lui Baesa în Israel, 15:32-16:7\r\ng. Domnia lui Zimri în Israel, 16:15-20\r\nh. Domnia lui Omri în Israel, 16:21-28\r\ni. Domnia lui Ahab În Israel, 16:29-22:40\r\n\r\nPROOROCUL ILIE\r\nIlie vesteste seceta, 17:1\r\nDumnezeu îl îngrijeste pe Ilie, 17:2-24\r\nIlie pe muntele Carmel, 18:1-46\r\nIlie fuge la Horeb, 19:1-18\r\nIlie H cheama pe Elisei, 19:19-21\r\n\r\nî Domnia lui Iosafat în Iuda, 22:41-50\r\nj. Domnia lui Ahazia în Israel, 22:51-53" },
    //2 ÎMPĂRAȚI
    { "2 ÎMPĂRAȚI", "Explicatia Cartii: 2 Imparati \r\n\r\n Începînd cu aceasta carte ne vom familiariza cu existenta marilor imperii antice: Asiria, Babilon si Persia. Ele vor începe sa exercite o influenta covîrsitoare asupra Israelului. Paradoxal, imperiile acestea sînt astazi o binecuvîntare pentru poporul Bibliei, caci recentele descoperiri arheologice au facut sa amuteasca critica arbitrara si rauvoitoare din deceniile trecute. Fiecare lopata de pamînt întoarsa în terenul stravechiilor imperii acrediteaza astazi Biblia ca o carte de istorie autentica.\r\nScopul scrierii acestei carti nu a fost numai acela de a înregistra vietile împaratilor din Iuda si Israel. Aceasta cronica a fost alcatuita pentru a arata tuturor cititorilor ca succesul oricarui împarat (si a oricarei natiuni luata ca atare) depinde de masura de atasament fata de Cuvîntul si voia Domnului. Neascultarea duce întotdeauna la decadenta si declin.\r\n2 Împarati este o continuare fireasca a cartii 1 Împarati cu care initial a constituit o singura cronica.\r\nCeea ce s-a spus despre 1 Împarati este valabil si pentru aceasta carte. Proorocul Ieremia apare ca cel mai probabil dintre toti autorii propusi. El a folosit însa informatii cuprinse în alte carti (1 Împ. 11:41; 14:19, 29). Ultimul capitol pare si el scris de altcineva care a stat în Babilon, nu de cineva care s-a dus cu unii din popor în Egipt si a murit acolo, cum este cazul cu Ieremia.\r\nÎn preajma anului 550 î.Cr. Nu înainte de distrugerea celui dintîi Templu.\r\n\r\nAceasta carte care debuteaza cu stramutarea lui Ilie la cer si se termina cu stramutarea poporului evreu în Babilon este cea mai tragica cronica de istorie din toata Biblia. Ba am putea spune ca este una dintre cele mai tragice înregistrari istorice din analele lumii. Poporul prin care Dumnezeu alesese sa raspîndeasca cunostinta despre Sine si despre planurile Sale pline de har s-a abatut din ce în ce mai mult de la ascultarea de Dumnezeu, s-a cufundat în idolatrie si în întunerecul înstrainarii de Creator. Cînd nelegiuirea lor si-a atins culmea, Dumnezeu a trimis asupra lor popoare straine puternice si lipsite de mila care i-au tîrît într-o neagra, umilitoare si chinuitoare robie.\r\n\r\nCartea 2 Împarati este cartea robiei. În capitolul 17 ni se prezinta ducerea Israelului (cele 10 semintii din împaratia de nord) în robia Asiriana. Aceste 10 semintii nu se vor întorce acasa, ci vor continua sa ramîna raspîndite printre neamurile lumii. Capitolul 25 ne aduce înaintea ochilor Ierusalimul înconjurat, cucerit si darîmat. Poporul împaratiei de sud (Iuda si Beniamin) este dus în robia babiloneana de unde nu se va mai întoarce decît o ramasita.\r\n\r\nCartea consemneaza de asemenea faptele profetului Elisei, continuatorul lucrarilor lui Ilie. Ca si predecesorul lui, Elisei are o viata plina de întîmplari supranaturale. Unele dintre cele mai cunoscute minuni savîrsite de el sînt: învierea fiului sunamitei (cap. 4), moartea în oala (cap.4), înmultirea celor 100 de pîini (cap.4), vindecarea lui Naaman, capetenia ostirii împaratului Siriei, (cap. 5), recuperarea securii unuia dintre fiii proorocilor (cap.6) si orbirea temporara a sirienilor (cap. 6).\r\n\r\nÎmpartirea cartii este neclara la prima vedere, dar un studiu atent va scoate la iveala faptul ca primele 10 capitole se ocupa predominant de împaratia de nord, implicînd împaratia de sud numai cînd evolutia acesteia afecteaza evenimentele din nord; partea dintre capitolele 11 si 27 este o istorie alternativa a evenimentelor din cele doua împaratii surori, iar de la capitolul 18 la 25 gasim numai evenimente din istoria lui Iuda.\r\n\r\nCu exceptia lui ªalum care a domnit numai o luna, toti cei 11 împarati ai lui Israel „au facut ce este rau înaintea Domnului\" (3:2, 3; 10:31, 32; 13:2, 3, 11; 14:24; 15:9, 18, 24, 28; 17:2). Ei au urmat calea lui Ieroboam, „cel care a facut pe Israel sa pacatuiasca\".\r\n\r\nÎmparatii din Iuda sînt evaluati dupa standardul stabilit de tatal lor David. Vedem aceasta în cazul lui Solomon (1 Împ. 11:6), Abiam (1 Împ. 15:3), Iosafat (2 Cronici 17:3), Amatia (2 Împ. 14:3), Ahaz (2 Împ. 16:2), Ezechia (2 Împ. 18:3) si Iosia (2 Împ. 22:2).\r\n\r\nPlecarea lui Israel în robie s-a facut în doua etape. Cele doua semintii si jumatate de dincolo de Iordan: Ruben, Gad si Manase au fost luate captive de Tiglat Pileser, împaratul Asiriei cu cîtiva ani înainte de restul împaratiei (2 Împarati 15:29; 1 Cronici 5:25-26). Celelalte semintii din împaratia de Nord au fost înrobite 13 ani mai tîrziu, cam prin anul 721 î.Cr. Tiglat Pileser murise între timp. Împaratul asirian Salmanasar a condus invazia si a asediat Samaria timp de trei ani (2 Împarati 17:3-6).\r\n\r\nCaderea împaratiei din sud, Iuda, s-a produs numai dupa alti 120 de ani de istorie. ªi ea a avut mai multe etape. Cea dintîi s-a petrecut în timpul domniei lui Ioiachim (2 Regi 23:36-24:1, 2; 2 Cronici 36:5-7). Atunci au fost duse la Babilon uneltele Casei Domnului împreuna cu cîtiva tineri din tara printre care si tînarul Daniel (Dan. 1: l -4). A doua deportare s-a produs opt ani mai tîrziu, la începutul domniei lui Ioiachin. Nebucadnetar l-a luat prizonier pe împarat si l-a stramutat la Babilon împreuna cu familia lui si cu 10 mii de oameni din floarea lui Israel. În Iuda n-au ramas decît oamenii saraci peste care Nebucadnetar l-a pus sa domneasca pe Zedechia, unchiul lui Ioiachin, întreg Ierusalimul a fost jefuit. Visteria Templului a fost luata si vasele de aur facute de Solomon la Templu au fost sparte (2 Împ. 24:8-17). A treia si cea din urma deportare în Babilon s-a produs în urma unei rascoale a lui Zedechia pe care Nebucadnetar a înabusit-o cu cruzime. Fiii lui Zedechia au fost înjunghiati înaintea lui, iar împaratului i-au scos ochii, l-au legat cu lanturi de arama si l-au dus la Babilon. Templul a fost distrus cu desavîrsire, cetatea Ierusalimului a fost transformata în ruine si tuturor caselor li s-a dat foc (2 Împ. 25:1-21). Asediul a fost îndelungat si a provocat celor din cetate suferinte de nedescris. Tragedia de proportii impresionante este descrisa în cartea sa de profetul Ieremia (Plîngeri 2:20; 4:3-10)\r\n\r\nPerioada împaratilor a fost si perioada de activitate a profetilor. Cartile lor vor fi analizate în sectiunea dedicata profetilor, dar este bine sa tinem minte ca Osea si Amos au profetit în Israel, în timp ce Obadia, Ioel, Isaia, Mica, Naum, Habacuc, tefania si Ieremia si-au desfasurat activitatea în Iuda.\r\n\r\nMesajul peste veacuri: Cine citeste paginile acestea de istorie a poporului Israel va trage repede concluzii pentru viata sa personala de umblare cu Domnul. Iata numai cîteva dintre ele: (1) Staruinta în neascultare este cea mai sigura cale spre dezastru. S-ar putea sa para pentru o vreme ca pedeapsa întîrzie, dar ea va veni cu siguranta. (2) Privilegiile acordate de Dumnezeu aduc cu ele responsabilitati corespunzatoare. Dumnezeu nu ne rasfata ca sa ne îngîmfe, ci ne binecuvînteaza ca sa ne îndemne sa staruim în împlinirea scopurilor pe care le-a fixat pentru existenta noastra. (3) Pedepsele trimise de Dumnezeu sînt proportionale si progresive. Scopul lui Dumnezeu este întotdeauna pedagogic. Cine nu raspunde la pedepse obisnuite îsi atrage singur pedepse mai mari si mai chinuitoare. (4) „Dumnezeu pedepseste pe cine iubeste si bate cu nuiaua pe orice fiu pe care îl primeste\" (Evrei 12:6). În plan uman si terestru se pare ca asistam la o catastrofa si la desfintarea poporului ales. În cer însa se socotise altfel. Providential, planul divin va continua cu Israelul. Cel mai mare profet din timpul declinului lui Israel declara: „El (Dumnezeu) nu va slabi, nici nu se va lasa, pîna va aseza dreptatea pe pamînt\" (Isaia 42:4). Cînd tronul de pe pamînt se prabuseste în pulbere, tronul din ceruri cîrmuieste stapîn pe furtuna. Robia babiloneana a venit ca o pedeapsa pentru decaderea idolatra a lui Israel. Evreii au trebuit sa învete departe de casa ce scump le este Dumnezeu si Legea Lui. În exil au aparut sinagogile si poporul a început sa citeasca si sa învete temeinic Cuvîntul. Robia a fost cuptorul în care Dumnezeu si-a curatit poporul. Raspînditi printre neamurile lumii, evreii au fost pastrati atunci si continua sa fie pastrati si astazi pentru împlinirea destinului lor istoric. Asirienii si babilonienii s-au pierdut în negura timpului, dar poporul evreu a continuat sa existe, în curînd va veni acel „Fiu al lui David\" care va prelua stapînirea în virtutea legamîntului Davidic si, din Ierusalimul restaurat, îsi va exercita autoritatea suprema asupra lumii.\r\n\r\nSCHItA CaRtII\r\n\r\nI. CRONICILE ÎMPaRatIEI DE NORD 1-10\r\na. Domnia lui Ahazia (853-852), 7:7-3\r\n\r\nILIE ªl ELISEI\r\nÎnaltarea lui Ilie la cer, 2:1-11\r\nElisei îsi începe slujirea, 2:12-14\r\nConfirmarea rapirii lui Ilie, 2:10-18\r\nVindecarea apelor de la Ierihon, 2:19-22\r\nPedepsirea copiilor, 2:23-25\r\n\r\nb. Domnia lui Ioram (852-841), 3:1-27\r\n\r\nFAPELE LUI ELISEI\r\nÎnmultirea untdelemnului, 4:1-7\r\nFiul Sunamitei, 4:8-37\r\nMoartea în oala, 4:38-41\r\nÎnmultirea pîinilor, 4:42-44\r\nVindecarea lui Naaman, 5:1-27\r\nSecurea care pluteste pe apa, 6:1-7\r\nElisei biruieste sirienii, 6:8-8:6\r\nElisei profeteste în Damasc, 6:7-15\r\n\r\nc. Domnia lui Ioram (852-841), 8:16-24\r\nd. Domnia lui Ahazia în Iuda (841), 8:25-29\r\ne. Domnia lui Iehu în Israel (841-814), 9:1-10:36\r\n\r\nII. CRONICa ALTERNATIVa 11-27\r\na. Domnia Ataliei în Iuda (841-835), 11:1-16\r\nb. Domnia lui Ioas în Iuda (835-796), 11:17-12:21\r\nc. Domnia lui Ioahaz în Israel (814-798), 13:1-9\r\nd. Domnia lui Ioas în Israel (798-782), 13:10-25\r\n\r\nMoartea lui Elisei 13:14-21\r\ne. Domnia lui Amatia (796-767), 14:1-22\r\nf. Domnia lui Ieroboam II în Israel (794-753), 14:23-29\r\ng. Domnia lui Azaria în Iuda (790-739), 15:1-7\r\nh. Domnia lui Zaharia în Israel (753), 15:8-12\r\ni. Domnia lui ªalum în Israel (752), 15:13-15\r\nî. Domnia lui Menahem în Israel (752- 742), 15:7 6-22\r\nj. Domnia lui Pecahia în Israel (742-740), 15:23-26\r\nk. Domnia lui Pecah în Israel (752-732), 15:27-31\r\nl. Domnia lui Iotam în Iuda (750-731), 15:32-38\r\nm. Domnia lui Ahaz în Iuda (731-715), 16:1-20\r\nn. Domnia lui Osea în Israel (732-722), 17:1-41\r\nSAMARITENII\r\n\r\nIII. CRONICILE ÎMPaRatIEI DE SUD 18-25\r\n\r\na. Domnia lui Ezechia (715-686), 18:1-20:21\r\nReforma religioasa, 18:1-12\r\nBiruinta sub doua asedii, 18:1-20:21\r\nBoala si însanatosirea lui, 20:1-11\r\nLaudarosenia lui vinovata, 20:12-21\r\n\r\nb. Domnia lui Manase (695-642), 21:1-18\r\nc. Domnia lui Amon (642-640), 21:19-26\r\nd. Domnia lui Iosia (640-609), 22:1-23:30\r\ne. Domnia lui Ioahaz (609), 23:31-33\r\nf. Domnia lui Ioiachim (609-597), 23:34-14:7\r\ng. Domnia lui Ioiachin (597), 24:8-16\r\nh. Domnia lui Zedechia (597-586), 24:17-25:21\r\nDarîmarea Ierusalimului si a Templului\r\n\r\ni. Guvernatorul Ghedalia (586), 25:22-26\r\nî. Eliberarea lui Ioiachin în Babilon, 25:27-30" },
    //1 CRONICI
    { "1 CRONICI", "Explicatia Cartii: 1 Cronici \r\n\r\nCartile Cronicilor descriu în mare aceleasi evenimente ca si cartile lui Samuel si ale Împaratilor. Atitudinea celui care le scrie este însa alta si scopul cu care sînt scrise este total diferit. 1 si 2 Samuel si 1 si 2 Împarati sînt carti cu un pronuntat caracter biografic, Cronicile au un caracter statistic. Primele sînt informative, cele din urma sînt formative, 1 si 2 Cronici sînt carti de educare a noii generatii de evrei nascuti în captivitatea babiloneana.\r\nCartea poarta aceasta numire din timpul lui Ieronim, care a tradus textul Bibliei ebraice în limba latina între anii 385-405 d.Cr. Lucrarea lui poarta numele de: „Vulgata\" (de la „vulg\" care înseamna popor). Titlul complet în limba latina este: „Cronicorum Liber\" - Cartea Cronicilor. Cartile cronicilor nu sînt carti de repetitii inutile ale evenimentelor relatate deja în cartile istorice ale Bibliei si nici culegeri de pasaje neimportante uitate de alti autori ai Scripturilor. Ele sînt compendii de istorie destinate evreilor din semintiile împaratiei lui Iuda care s-au întors din robia Babiloniana.\r\nDesi numele autorului este necunoscut, parerile celor mai multi înclina spre carturarul Ezra, cel care a fost responsabil pentru marea lucrare de reconstruire spirituala a evreilor dupa revenirea din Babilon.\r\nCeea ce este cert si evidentiat expres în textul Cronicilor este ca autorul s-a folosit de un mare numar de lucrari istorice existente. Iata o lista a lor:\r\n\r\n1. Cartea împaratilor lui Israel si a lui Iuda (2 Cronici 27:7)\r\n2. Un „Midras\" (comentariu) al lucrarii de mai sus (2 Cronici 24:27)\r\n3. Cartea lui Samuel vazatorul (1 Cronici 29:29)\r\n4. Cartea proorocului Natan (1 Cronici 29:29)\r\n5. Cartea proorocului Gad (1 Cronici 29:29)\r\n6. Proorocia lui Ahia din ªilo (2 Cronici 9:29)\r\n7. Descoperirile proorocului Ieedo (2 Cronici 9:29)\r\n8. Cartea proorocului ªemaia (2 Cronici 12:15)\r\n9. Cartea de genealogii a proorocului Ido (2 Cronici 12:15)\r\n10. Istoria alcatuita de Iehu, fiul lui Hanani (2 Cronici 20:34)\r\n11. Cartea proorocului Isaia (26:22)\r\n12. Vedenia proorocului Isaia (2 Cronici 32:32)\r\n13. Cartea lui Hozai (2 Cronici 33:19)\r\n\r\nMultitudinea aceasta de izvoare istorice este mult mai importanta decît pare la prima vedere. Ea dovedeste ca: (a) autorul a cercetat mult înainte de a se apuca de lucru, (b) autorul nu se sfieste sa citeze titlurile surselor lui, ceea ce dovedeste ca ele erau materiale demne de toata cinstea si încrederea, si ca (c) colectarea datelor istorice a fost o îndeletnicire constanta de-a lungul istoriei lui Israel, iar aceasta ne da si mai multa încredere în veridicitatea înregistrarilor istorice pe care se bazeaza învatatura Bibliei.\r\nDe la sfîrsitul cartilor 1 si 2 Împarati au trecut 70 de ani de robie. Profetiile rostite de Ieremia s-au împlinit. Captivitatea a fost cuptorul în care Dumnezeu si-a vindecat poporul de boala idolatriei. Printre straini, evreii au fost învatati ca este mai bine sa fie aruncati în cuptorul cu foc decît sa-si plece genunchii în fata chipurilor idolatre (Daniel 3). Mai marii lor le-au aratat ca este mai bine sa fie aruncati în groapa cu lei decît sa-si întrerupa viata de rugaciune catre Iehova (Daniel 6).\r\n\r\nÎn 70 de ani murise însa o generatie. Avem de-a face acum cu un nou Israel, nascut departe de tara, departe de Ierusalim si departe de Templu. Ei fusesera crescuti printre straini, fara o viata de ceremonii religioase si fara cunostinta destinului lor providential. În anul 537 î.Cr. decretul împaratului persan Cir da dreptul evreilor sa se întoarce acasa la Ierusalim si în Iuda. Tinerii nascuti în captivitate si-au pierdut însa legatura fireasca de continuitate cu generatia dinainte de captivitate. Cineva trebuie sa înceapa o lucrarea de educatie prin care sa fie reînnodat firul lor de legatura cu trecutul. Cineva trebuie sa alcatuiasca o punte de trecere între ceea ce a fost, ceea ce este si ceea ce trebuie sa urmeze. Cartile 1 si 2 Cronici fac exact aceasta lucrare.\r\nAutorul cronicilor nu este un pasionat de istorie, ci un iubitor de oameni. Dorinta lui nu este aceea de a transmite date reci, ci de a da informatii care sa determine un anumit fel de vietuire. Cronicile scrise de el sînt scrise cu o anumita atitudine si cu un anumit scop. Atitudinea este grija pentru reînodarea istoriei providentiale a lui Israel, iar scopul este educatia religioasa a noii generatii care s-a întors în tara. (SELECTARE, ADaUGARE, INTERPRETARE)\r\n\r\nEra important ca cineva sa faca natiunea sa-si înteleaga trecutul, prezentul si viitorul din punctul de vedere a lui Dumnezeu. Acum, cînd tronul lui David nu mai era punctul de raliere si de directionare a poporului, trebuia vazut ce mai ramasese din fiinta nationala. Trei lucruri au fost scoase în evidenta:\r\n\r\nl. Genealogiile ca o revenire în matca tiparului divin (1 Cron. 1-9). Fara nici o îndoiala ca din cauza robiei, multe familii din Israel îsi pierdusera evidentele genealogice (vezi Ezra 2:59). Ori aceasta evidenta era foarte importanta pentru reasezarea natiunii în teritoriile tarii.\r\n\r\n2. O recapitulare a istoriei, mai ales a istoriei recente, din care sa fie deduse cauzele tragediei robiei. Aceasta lectie trebuia învatata bine pentru ca sa nu mai fie repetata niciodata (1 Cron. 10-29; 2 Cronici 1-36)\r\n\r\n3. O refacere a vietii religioase centrata în jurul Templului de la Ierusalim. Poporul trebuia însufletit pentru reconstruirea Templului si pentru reluarea vietii de închinaciune. În absenta tronului împaratesc al lui David, centru de greutate al Israelului trebuia asezat acolo unde se afla simbolic tronul împaratesc al cerului.\r\n\r\nPentru a-si realiza toate aceste trei deziderate, autorul Cronicilor face o munca de selectare, de adaugare si de interpretare a datelor pe care le gaseste în izvoarele istorice. El polarizeaza evenimentele si personajele istoriei trecute în asa fel încît citirea Cronicilor lui sa devina o lectie pedagogica pentru pregatirea viitorului. Anumite secvente de istorie sînt lasate afara intentionat: despre David, de exemplu, nu se aminteaste nimic despre ratacirea lui sentimentala cu Bat-Seba si uciderea lui Urie si nici despre cîntarea lui de jale dupa Saul si Ionatan sau despre rascoala tînarului Absalom. Alte secvente de istorie sînt tratate mai amanuntit si chiar completate cu informatii pe care nu le gasim în cartile lui Samuel sau în cele ale împaratilor: pregatirile intense facute de David pentru construirea Templului (1 Cron. 22), pregatirea si repartizarea preotilor, levitilor (1 Cron. 23-24), si rînduirea cîntaretilor, usierilor si vistierilor de la Templu (1 Cron. 25-26). Domniile anumitor împarati sînt scoase mai mult în evidenta si tratate mai pe larg deoarece ele constituie pilde de reforma religioasa si de revitalizare spirituala: Asa, Iosafat, Ioas, Ezechia si Iosia (toti în 2 Cronici).\r\nMesajul imediat pentru prima generatie de cititori a fost întreit: (1) natiunea trebuie sa se reaseze cît mai degraba si cît mai temeinic în tiparul si prerogativele de popor mesianic al legamîntului, (2) natiunea trebuie sa învete din catastrofa suferita de curînd si sa nu mai repete greselile care au provocat-o si (3) natiunea trebuie sa-si reia partasia de la Templu cu Iehova. Se vede imediat ca autorul subliniaza cu precadere o idee care se poate vedea din întreaga istorie a lui Israel si anume ca, spre deosebire de alte neamuri, evreii nu au nici un rost pe pamînt fara Iehova. Valoarea si specificul lor este în viata lor de închinaciune. Destinul lor este acela de a-L face cunoascut lumii pe Dumnezeu Creatorul lumii care vrea sa refaca legatura cu creatura cazuta. Israel este instrumentul care trebuie sa dea lumii doua lucruri: cunostinta despre Dumnezeu si persoana lui Mesia.\r\n\r\nMesajul pe care ni-l trimite noua cartea nu este prea deosebit de acesta. ªi noi trebuie sa întelegem ca legatura noastra cu Dumnezeu este singurul lucru care conteaza în aceasta existenta. Ceea ce ni se întîmpla în viata, evenimentele bucuriilor si durerilor noastre, nu trebuiesc privite decît drept niste lectii prin care Dumnezeu se straduieste sa ne învete acest adevar de referinta.\r\nI. GENEALOGIILE LUI ISRAEL l - 9\r\na. De la Adam la Avraam, 1:1-27\r\nb. De la Avraam la Iacov, 1:28-54\r\nc. De la Iacov la David, 2:1-55\r\nd. De la David la robie, 3:1 -24\r\n\r\nGENEALOGIA CELOR 12 SEMINtII\r\n1. Iuda, 4:1-23\r\n2. Simeon, 4:24-34\r\n3. Ruben, 5:1-10\r\n4. Gad, 5:11-22\r\n5. Manase, 5:23-26\r\n6. Levi, 6:1-18\r\n7. Isahar, 7:1-5\r\n8. Beniamin, 7:6-12\r\n9. Neftali, 7:13\r\n10. Manase, 7:14-19\r\n11. Efraim, 7:20-29\r\n12. Aser, 7:30-40\r\n13. Beniamin, 8:1-40\r\n\r\ne. Locuitorii Ierusalimului, 9:1-34\r\nf. Familia lui Saul, 9:35-44\r\n\r\nII. DOMNIA LUI DAVID 10-29\r\n\r\nUNSUL DOMNULUI\r\na. Moartea lui Saul, 10:1-14\r\nb. Ridicarea lui David, 11:1-3\r\nc. Cucerirea Ierusalimului, 11:4-9\r\nd. Vitejii lui David, 11:10-12:40\r\n\r\nCHIVOTUL DOMNULUI\r\na. David aduce chivotul la Chidon, 13:1-14\r\nb. Biruinta asupra filistenilor, 14:1-17\r\nc. David aduce chivotul la Ierusalim, 15:1-29\r\nd. Orînduieli privitoare la slujba, 16:1-43\r\n\r\nLEGaMÎNTUL DOMNULUI\r\na. David vrea sa zideasca Templul 17:1-27\r\nb. Razboaiele lui David, 18:1-20:8\r\nc. Numaratoarea vinovata, ciuma, 21:1-30\r\n\r\nTEMPLUL DOMNULUI\r\na. Pregatirile pentru Templu, 22:1-23:1\r\nb. Organizarea levitilor, 23:2-26:32\r\nc. Capeteniile ostirii, 27:1-34\r\nd. Testamentul lui David, 28:1-21\r\ne. Daruri pentru Templu, 29:1-9\r\nf. Rugaciunea lui David, 29:10-19\r\ng. Ungerea lui Solomon, 29:20-25\r\nh. Moartea lui David, 29:26-30" },
    //2 CRONICI
    { "2 CRONICI", "Explicatia Cartii: 2 Cronici\r\n\r\nCartea 2 Cronici este o dizertatie istorica pe tema: Soarta unei natiuni depinde de atitudinea ei fata de Domnul. Cazul studiat este: „Împaratia lui Iuda\". Cartea este continuarea lucrarii începute în cea dintîi carte a Cronicilor.\r\nContinutul Cronicilor se termina brusc cu o fraza întrerupta la jumatate (2 Cron. 36:23). Continuarea ei se poate gasi în primele versete ale cartii lui Ezra, carturarul (Ezra 1:2-4). Concluzia fireasca ce se poate trage din aceasta „coincidenta\" este ca ambele carti au acelasi autor, desi numai cea din urma îi poarta numele.\r\nDupa darea edictului lui Cir care a permis reîntoarcerea evreilor în Israel, probabil între anii 450-425 î.Cr. (1 Cron. 3:16-24; 9:1)\r\n2 Cronici este o carte tragica cu un început glorios, dar cu un final teribil. Primele 9 capitole ne redau cei 40 de ani de domnie ai lui Solomon. Restul capitolelor se concentreaza asupra istoriei împaratiei lui Iuda pîna în vremea robiei babiloniene. Nu este cazul sa zabovim aici asupra fiecaruia dintre cei 20 de împarati care s-au succedat la tronul Ierusalimului. Este suficient sa spunem ca ei sînt analizati de autor prin prisma relatiei personale pe care au avut-o ei cu Domnul. Ascensiunea sau declinul împaratiei lui Israel nu s-au datorat factorilor economici sau politici, ci starii lor spirituale în legatura lor cu Dumnezeu: „ªi în timpul cînd a cautat pe Domnul, Dumnezeu l-a facut sa propaseasca\" (2 Cron. 26:5), „Iotam a ajuns puternic, pentru ca si-a urmat necurmat caile înaintea Domnului, Dumnezeului sau\" (2 Cron. 27:6), „Ahaz... El n-a facut ce este bine înaintea Domnului... Domnul, Dumnezeu l-a dat în mîinile împaratului Siriei\" (2 Cron. 28:1-5).\r\n\r\nÎntr-un sens foarte larg, Israelul, ca popor, este prezentat în cartile Cronicilor ca si „casa lui Iehova\". Dumnezeu îi spune lui David: „Domnul îti va zidi o casa\" (1 Cron. 17:10). Tot ceea ce se întîmpla în popor este în functie de atitudinea lor fata de Dumnezeu si ca urmare a binecuvântarilor sau pedepselor pe care le trimite El asupra natiunii: „Caci voi cinsti pe cine Ma cinsteste, dar cei ce Ma dispretuiesc, vor fi dispretuiti\" (1 Sam. 2:30).\r\n\r\nCuvinte cheie si teme caracteristice: Toata actiunea istorica poate fi grupata pe doua coloane în fruntea carora se poate scrie: „ascultare\" si „neascultare\". Cronicile sînt lectii pedagogice pentru copiii lui Dumnezeu din toate timpurile.\r\n\r\nMesajul peste veacuri: A fost adevarat atunci si este adevarat si acum; este adevarat pentru neamuri întregi si este adevarat pentru fiecare om în parte: Cea dintîi si cea mai de capetenie datorie a oricarui tron este o buna relatie cu templul. Spusa în cuvintele nemuritoare ale Domnului Isus aceasta axioma suna asa: „Cautati mai întîi împaratia lui Dumnezeu si neprihanirea Lui, si toate celelalte lucruri vi se vor da pe \"deasupra\" (Matei 6:33).\r\nI. DOMNIA LUI SOLOMON I- 9\r\na. Împaratul Solomon, 1:1-17\r\n\r\nTEMPLUL LUI SOLOMON\r\nPregatiri pentru Templu, 2:1-18\r\nConstruirea Templului, 3:1-4:22\r\nDedicarea Templului, 5:1-7:22\r\n\r\nb. Faima lui Solomon, 8:1-9:29\r\nc. Moartea lui Solomon, 9:13-28\r\n\r\nII. ÎMPaRAtII LUI IUDA, 10-36\r\n\r\na. ROBOAM\r\nRoboan produce dezbinarea, 10:1-19\r\nRoboam slujeste Domnului, 11:1-23\r\nRoboam paraseste pe Domnul, 12:1-16\r\n\r\nb. Abia, 13:1-22\r\nc. Asa, 14:1-16:14\r\n\r\nd. IOSAFAT\r\nReforma religioasa, 17:1-19\r\nAlianta cu Ahab, 18:1-19:3\r\nAlte reforme, 19:4-11\r\nBiruinta asupra lui Moab si Amon, 20:1-30\r\nUltimele lui zile, 20:31-37\r\n\r\ne. Ioram, 21:1-20\r\nf. Ahazia, 22:1-9\r\ng.Atalia, 22:10-23:15\r\nh. Ioas, 23:16-24:27\r\ni. Amatia, 25:1-28\r\nf. Ozia, 26:1-23\r\nj. Iotam, 27:1-9\r\nk. Ahaz, 28:1-27\r\n\r\nl. EZECHIA\r\nTrezirea spirituala, 29:1-31:21\r\nBiruinta asupra asirienilor, 32:1-23\r\nUltimele lui zile, 32:24-33\r\n\r\nm. Manase, 33:1-20\r\nn. Amon, 33:21-25\r\no. Iosia, 34:1-35:27\r\np. loahaz, 36:1-4\r\nr. Ioiachim, 36:5-8\r\ns. Ioiachin, 36:9-10\r\ns. Zedechia, 36:11-21\r\n\r\nIII. DECRETUL LUI CIRUS 36:22-23" },
    //EZRA
    { "EZRA", "Explicatia Cartii: Ezra\r\n\r\nCele trei carti mai mici care ne stau acum în fata - Ezra, Neemia si Estera - încheie seria celor 17 carti istorice care alcatuiesc prima sectiune a Vechiului Testament. Aceste trei carti trebuiesc studiate împreuna deoarece ele înregistreaza lucrarea lui Dumnezeu cu poporul Israel dupa întoarcerea evreilor din robia Babiloniana. Ezra si Neemia se ocupa de „ramasita\" care s-a întors în Ierusalim si în Iudea, iar Estera explica ce s-a întîmplat cu poporul care a preferat sa ramîna prin tarile în care fusesera dusi. Pentru o mai buna întelegere a acestor carti care marcheaza sfîrsitul cronicilor istorice este bine sa citim si ultimele trei carti din sectiunea cartilor profetice: Hagai, Zaharia si Maleahi, caci ei au fost profetii ridicati de Dumnezeu în Israel în perioada de dupa robie. Daca aceasta carte ar fi fost numita nu dupa numele autorului ei, ci dupa continutul pe care îl are, ea ar fi putut fi intitulata: „Cartea restaurarii\", „Cartea repatrierii\" sau „Cronica ramasitei lui Iuda\". Daca ar fi fost numita dupa numele eroilor ei principali, ar fi trebuit sa i se spuna: „Cartea lui Zorobabel si a lui Ezra\". Titlul ei este însa „Ezra\" si aceasta subliniaza si mai mult rolul jucat de aceasta extrordinara personalitate în lucrarea pe care Dumnezeu a facut-o pentru reasezarea Israelului în tiparul divin. Ezra face parte din marele triumvirat al carturarilor responsabili pentru scrierea Vechiului Testament: Moise, Samuel si Ezra. Traditia evreiasca, prin Talmud, îl considera pe Ezra unul dintre cei mai proeminenti lideri din istoria lui Israel. Lui îi sînt atribuite cinci lucrari capitale pentru identitatea evreilor ca neam: (1) înfiintarea în Babilon a unei institutii cu numele de: „Sinagoga cea mare\". Aceasta scoala rabinica era alcatuita dintr-un numar de 120 de membrii, toti urmasi ai profetilor si întemeietori ai gruparii „carturarilor\" care au pastrat si transmis apoi scrierile sfinte evreiesti; (2) alcatuirea „canonului\" Vechi Testamental. Sub acest nume este descrisa „culegerea de Scripturi sfinte evreiesti recunoscute de toti evreii ca avînd autoritate dumnezeiasca\". Ezra le-a asezat pe toate în trei categorii distincte numindu-le: Legea, Profetii si Scrierile; (3) trecerea de la scrierea antica ebraica la scrierea evreiasca cu caractere grafice asiriene; (4) alcatuirea Cronicilor ca un rezumat al istoriei evreiesti si scrierea cartilor Ezra si Neemia; (5) înfintarea sinagogilor locale ca locuri în care poporul sa citeasca si sa învete perceptele religiei evreiesti.\r\n\r\nFaptul ca Ezra este socotit autorul originalului ebraic al cartilor 1 si 2 Cronici, Ezra si Neemia nu înseamna ca el nu a folosit fragmente din scrierile altor autori sau ca portiunea autobiografica din cartea Neemia nu ar fi putut fi scrisa chiar de Neemia însusi.\r\nCartea se ocupa cu unul dintre cele mai importante evenimente din istoria Israelului: întoarcerea din robie. Aceasta întoarcere nu s-a facut într-o singura etapa, ci în mai multe. Tot asa si pentru consemnarea evenimentelor putem presupune ca exista mai multe date ale scrierii. Refacerea vietii religioase a Israelului a durat o perioada de aproximativ o suta de ani. În acest interval de timp au existat doua perioade mai mici de o exceptionala importanta:\r\n\r\n1. Cei 20 de ani (537-517 î.Cr.) scursi între darea decretului lui Cir-persanul si pîna într-al saselea an al domniei lui Dariu-medul, în care, sub conducerea lui Zorobabel ca guvernator si Iosua ca preot, poporul a rezidit Templul. Pentru cunoastera evenimentelor din aceasta perioada cititi Ezra 1-6, Zaharia si Hagai, ultimele doua capitole din cartea 1 Cronici, primul capitol din cartea 2 Cronici, Psalmii 126 si 137, precum si referintele despre împaratul persan Cir din cartea profetului Isaia (Isaia 44:23 pîna la 45:8).\r\n\r\n2. Cei 25 de ani (458-433 î.Cr.) în care Neemia, ca guvernator, Ezra, ca preot au condus lucrarea de refacere a zidurilor Ierusalimului, în vremea aceea a activat profetul Maleahi.\r\n\r\nÎn cartea Ezra ne întîlnim cu amîndoua aceste perioade, în Neemia gasim numai cea de a doua perioada.\r\nEvreii au fost dusi în robie mai întîi de asirieni (1 Împ. 17) si apoi de babilonieni (2 Împ. 25). Dupa 70 de ani de robie, împaratul persan Cir le-a dat dreptul de a se întoarce în tara lor. Babilonul cazuse între timp sub cucerirea medo-persanilor (Octombrie 539 î.Cr.). Cele zece semintii luate în robie de asirieni nu s-au mai întors în Israel, în anul 537 î.Cr. s-au întors în Israel primii evrei veniti din robie. În anul 516 î.Cr. a fost încheiata recladirea Templului, în anul 479 î.Cr. Estera a devenit împarateasa Persiei (ca sotie a lui Xerxes). În anul 458 î.Cr. Ezra a condus cel de al doilea convoi de evrei care s-au întors din robie, în anul 445 î.Cr. Neemia a reconstruit zidurile Ierusalimului.\r\nEzra ne aduce înaintea ochilor cel de al doilea „Exod\" din istoria evreilor. Primul avusese loc în Egipt sub conducerea lui Moise. Cel de al doilea are loc acum din Babilon sub conducerea lui Ezra. Ca si Moise, Ezra este un om al cartii, cu o educatie aleasa si destinat sa alcatuiasca culegerea de cartii sfinte ale neamului.\r\n\r\nAvem toate motivele sa credem ca la început, 1 si 2 Cronici, Ezra si Neemia au format o singura scriere istorica. Evreii si crestinii din primele veacuri l-au considerat pe Ezra drept autor al acestor compilatii.\r\n\r\nNumarul celor întorsi din Babilon cu ocazia plecarii primului convoi a fost de aproximativ 50.000 de oameni. Este un numar relativ mic fata de totalul evreilor care plecasera în robie si cu mult mai mic decît numarul total al evreilor aflati în exil la aceea ora. Cei ce s-au întors au alcatuit într-adevar numai o „ramasita\". Ceilalti au preferat sa uite necazurile stramosilor lor si s-au hotarît sa ramîna între neamurile printre care fusesera stramutati. Se poate ca multi dintre ei sa fi reusit între timp sa-si faca un rost si o situatie materiala prospera. Perspectiva întoarcerii în saracia Israelului pentru a-i rezidi darîmaturile nu a parut prea atragatoare multora dintre ei. Întoarcerea celor 50.000 si faptele lor sînt descrise în primele 6 capitole ale cartii Ezra. Celelalte capitole, (7-10), descriu întoarcerea lui Ezra la Ierusalim si lucrarea lui de refacere si reforma spirituala a poporului.\r\n\r\nCuvinte cheie si teme caracteristice: întreaga carte graviteaza în jurul declaratiei urmatoare: „Caci Ezra îsi pusese inima sa adînceasca si sa împlineasca Legea Domnului, si sa învete pe oameni în mijlocul lui Israel legile si poruncile\" (Ezra 7:10).\r\nCartea lui Ezra este un tipar pentru miscarile de trezire spirituala. Continutul ei ne arata 6 pasi spre o stare dupa voia lui Dumnezeu:\r\n\r\n1. Întoarcerea în tara (cap.1 si 2) - revenirea la stravechea vatra a credintei.\r\n2. Recladirea altarului (cap. 3:1-6 - reluarea unei vieti de rugaciune.\r\n3. Recladirea Templului (3:8-13) - reconsiderarea atitudinii fata de Dumnezeu.\r\n4. Înfruntarea dusmanilor (cap.4) - respingerea ispitelor si tentatiilor care ne-ar putea opri din drum.\r\n5. Îndemn profetic (5:1-6:14) - supunerea fata de autoritatea Cuvîntului lui Dumnezeu.\r\n6. Terminarea Templului (6:15-22) - perseverenta în ascultarea si cinstirea Domnului.\r\n\r\nO astfel de refacere spirituala trebuie sa se produca din cînd în cînd în fiecare dintre noi. Ea este necesara mai ales atunci cînd crestinismul nostru s-a poticnit undeva în trecut si ajungem sa ne dam seama ca traim \"robia\" sub apasarea dusmanilor sufletului nostru. Din cînd în cînd este bine sa se ridice cîte un Pavel care sa preia mesajul lui Ezra si sa ne întrebe: „Oh, Galateni nechibzuiti! Cine v-a, fermecat pe voi, înaintea ochilor carora a fost zugravit Isus Cristos ca rastignit?\" „Voi alergati bine; cine v-a taiat calea ca sa n-ascultati de adevar?\" (Galat. 3:1; 5:7)\r\n\r\nMesajul despre Dumnezeu pe care îl gasim în aceasta carte poate fi parafrazat cu un text din Ieremia: „Caci Domnul nu leapada pentru totdeauna. Ci, cînd mîhneste pe cineva, Se îndura iarasi de el, dupa îndurarea Lui cea mare\" (Plîngerile lui Ieremia 3:31-32).\r\nSCHITA CARTII\r\n\r\nI. ÎNTOARCEREA LUI ZOROBABEL 1-6\r\nDecretul lui Cir, 1:1-11\r\n\r\nNUMaRaTOAREA POPORULUI\r\nConducatorii, 2:1-2\r\nFamiliile, 2:3-20\r\nCetatile, 2:21-35\r\nPreotii, 2:36-39\r\nLevitii, 2:40-42\r\nSlujitorii, 2:43-54\r\nRobii lui Solomon, 2:55-58\r\nDescendenti neconfirmati, 2:59-63\r\nCifra totala, 2:64-70\r\n\r\nRECONSTRUIREA TEMPLULUI\r\nRezidirea altarului, 3:1-3\r\nSarbatoarea Corturilor, 3:4-6\r\nRefacerea temeliilor Templului, 3:7-13\r\n\r\nAMÎNAREA ZIDIRII TEMPLULUI\r\nSamaritenii vor sa participe, 4:1 -3\r\nCompromisul refuzat, 4:4-5\r\nConsecinta, 4:6-24\r\nZidirea Templului începe iarasi, 5:1-17\r\nDecretul lui Dariu, 6:1-12\r\nTerminarea zidirii, 6:13-15\r\nSfintirea Templului, 6:16-22\r\n\r\nII. ÎNTOARCEREA LUI EZRA 7-10\r\nÎntoarcerea la Ierusalim, 7:1-8:36\r\n\r\nTREZIREA SPIRITUALa DE LA IERUSALIM\r\nStarea poporului, 9:1-4\r\nMarturisirea lui Ezra, 9:5-15\r\nLegamîntul cu poporul, 10:1-8\r\nCuratirea poporului, 10:9-44\r\n\r\n\r\n\r\n\r\n\r\n\r\nÎMPaRAtI si \r\n\r\n\r\n\r\n DATE\r\n\r\n\r\n\r\nCAPITOLE ÎN CARTEA EZRA \r\n\r\n\r\n\r\n TEXTE corespondente\r\n\r\n\r\n\r\nCirus\r\n538-530\r\ncapitolele 1-6\r\n \r\n\r\n\r\nCambises\r\n530-522\r\n(desi nu sînt mentionate numele \r\nlui Cambises si Smerdis)\r\nHagai si\r\nZaharia\r\n\r\n\r\nSmerdis\r\n522\r\n \r\n \r\n\r\n\r\nDarius\r\n521-486\r\n \r\n \r\n\r\n\r\nXerxes I\r\n486-465\r\n4:6\r\nEstera\r\n\r\n\r\nArtaxerxe I\r\n464-423\r\n4:7-23 si cap. 7-10\r\nMaleahi\r\n\r\n\r\nDarius II\r\n423-404\r\n \r\nNeemia\r\n\r\n" },
    //NEEMIA
    { "NEEMIA", "Explicatia Cartii: Neemia\r\n\r\nExilul babilonian a semnat actul de deces al limbii ebraice vechi. Poporul nascut în robie si-a însusit limba cu caractere aramaice. De aici s-au tras o seama de dificultati legate de educatia religioasa. Cu unele dintre ele ne vom întîlni în cartea Neemia. Cartea poarta numele personajului ei principal. Neemia a fost unul dintre evreii ramasi în „diaspora\" (raspîndire) atunci cînd împaratul Cir a dat evreilor decretul de întoarcere în tara promisa. El a ocupat functia de paharnic al împaratului Artarxerxes, o pozitie de onoare si de mare influenta asupra deciziilor imperiale. Acest Artaxerxes a avut-o ca mama vitrega pe Estera, printesa iudeica care traia fara îndoiala în perioada de timp acoperita de aceasta carte. Probabil ca influenta ei l-a ajutat pe Neemia sa ocupe aceasta functie înalta în ierarhia de la curtea împarateasca.\r\nFara nici o îndoiala ca partile scrise la persoana a-I-a sînt scrise de Neemia însusi. (De la cap. 1 la 7 si de la 12:27 la 13:31 care marcheaza sfîrsitul cartii). Restul poate fi usor contributia carturarului Ezra, asa cum am aratat deja în introducerea facuta la cartea care-i poarta numele.\r\nCu aproximatie, anul 430 î.Cr. Prima venire a lui Neemia la Ierusalim „în luna Nisan, anul al douazecelea al împaratului Artaxerxe\", a fost stabilita de Societatea Astronomica Britanica ca 14 Martie 445 î.Cr. A doua venire a lui Neemia la Ierusalim a avut loc în „al treizeci si doilea an al împaratului Artaxerxe\", ceea ce înseamna anul 432 î.Cr.\r\nAsa cum am aratat, Neemia a sosit pentru prima data la Ierusalim în anul 445 î.Cr. „Ramasita\" iudeilor întorsi din robie se afla deja în tara de aproximativ 90 de ani. Zorobabel si generatia lui murisera între timp. În locul lor traia acum o alta generatie, în cei 90 de ani de vietuire în Israel, evreii reconstruisera Templul, e drept cu dimensiuni mult mai mici si mai modeste decît fosta constructie ridicata de Solomon. Întreaga lucrare durase numai patru ani, patru luni si zece zile (Hagai 1:15 si Ezra 6:15). Dupa alti 60 de ani venise în Ierusalim si Ezra însotit de 2.000-3.000 de oameni (Ezra 8:1-14). Conditiile morale si spirituale în care traiau evreii erau de plîns. Mai marii poporului, preotii, levitii si multi din popor contractasera casatorii mixte cu femeile din neamurile idolatre din jur. Chiar daca nu a produs o întoarcere totala spre idolatrie, aceasta conditie încuraja practicarea ei pe teritoriul Israelului. Mai mult, stricarea puritatii neamului punea în pericol însasi existenta lui distincta în istorie. Nu este de mirare ca Ezra s-a umplut de consternare si de mînie (Ezra 9:3-15).\r\n\r\nAcum, dupa alti 12 ani, se întoarce la Ierusalim Neemia. Situatia gasita de el nu este mai putin tragica. Zidurile cetatii erau înca în ruina, iar poporul era fara viziune si fara vlaga, departe de stralucirea care ar fi trebuit sa-i încununeze destinul mesianic.\r\nObiectivul vizitei lui Neemia a fost rezidirea zidurilor Ierusalimului. Am vazut cum cartea lui Ezra a fost împartita în doua sectiuni majore: refacerea Templului si refacerea vietii de închinaciune. Acum cartea lui Neemia se împarte similar în: refacerea zidurilor cetatii (Neemia 1-4) si restaurarea poporului în ascultare si împlinire a Legii (Neemia 7-13). Carturarul Ezra este ajutorul lui Neemia în lucrarea de citire si explicare a Legii catre popor (Neemia 8:1-18).\r\nNeemia este un exemplu de dedicatie si daruire. El a fost gata sa lase prestigiul si confortul de la curtea împaratului Artaxerxe pentru a participa la reconstructia materiala si la restaurarea religioasa a poporului sau. Neemia a realizat o lucrare monumentala într-un interval record de timp. Acest om a fost un geniu al administratiei si organizarii. Un studiu al metodelor folosite de el este un izvor de inspiratie pentru conducatorii din toate timpurile. El a fost un om al rugaciunii, al credintei, al curajului si al actiunii. Comentînd viata lui Neemia am putea spune: „Doamne, fa-ne spirituali în întregime, dar lasa-ne plini de naturalete si mai ales de spirit practic!\".\r\nI. RECONSTRUIREA ZIDULUI l - 6\r\na. Situata Ierusalimului, 1:1-3\r\nb. Reactia lui Neemia, 1:4\r\nc. Rugaciunea lui Neemia, 1:5-11\r\nd. Plecarea la Ierusalim, 2:1-10\r\ne. Îmbarbatarea poporului, 2:11-20\r\nf. Repartizarea sectoarelor, 3:1-32\r\n\r\ng. Rezistenta în fata atacurilor\r\nAtacul batjocorei, 4:1-6\r\nAtacul prin forta, 4:7-23\r\nAtacul nedreptati, 5:1-19\r\nAtacul compromisului, 6:1-4\r\nAtacul înselaciunii, 6:10-14\r\n\r\nh. Terminarea lucrarii, 6:15-7:4\r\n\r\nII. REFORMA RELIGIOASa 7-13\r\na. Reînregistrarea poporului, 7:5-73\r\nb. Recitirea Legii, 8:1-8\r\nc. Raspunsul poporului, 8:9-18\r\nd. Pocainta poporului, 9:1-38\r\ne. Refacerea legamîntului, 10:1-39\r\nf. Repopularea oraselor, 11:1-12:26\r\ng. Rededicarea zidurilor, 12:27-47\r\n\r\nh. Reforme religioase\r\nÎn relatiile cu ne-evreii, 13:1-3\r\nÎn relatiile cu preotimea, 13:4-14\r\nÎn relatia cu Sabatul, 13:15-22\r\nÎn relatiile de casatorie, 13:23-31" },
    //ESTERA
    { "ESTERA", "Explicatia Cartii: Estera\r\n\r\nCele trei mici carti de la sfîrsitul sectiunii rezervate canilor istorice: Ezra, Neemia si Estera ne dezvaluie lucrarea facuta de Dumnezeu cu evreii întorsi din robia babiloniana. Deosebirea dintre cartea Estera si celelalte doua este ca, în timp ce Ezra si Neemia descriu soarta celor întorsi în tara lui Israel, Estera descrie un eveniment petrecut în viata milioanelor de evrei ramasi rasfirati prin toate tarile imperiului. Împreuna cu cartea Rut, Estera este singura carte din Biblie care poarta drept titlu un nume de femeie. Aceasta eroina a neamului ei, fusese mai întîi numita Hadasa, dar numele ei evreiesc a fost schimbat pentru folosul celor de la curtea imperiala în „Estera\" care se traduce prin „steaua rasaritului\".\r\nNimeni nu stie cu certitudine cine este autorul acestei remarcabile carti.\r\nNu este nici ea cunoscuta. Actiunea cartii se petrece însa ca timp undeva între capitolele 6 si 7 ale cartii lui Ezra.\r\nProbabil ca ati auzit de Xerxes, unul dintre cei mai renumiti împarati ai antichitatii, care a pornit cu o expeditie militara împotriva Greciei si a carei flota a fost înfrînta în batalia de la Salamina (480 î.Cr.). A fost una dintre cele mai importante batalii navale din istorie. Xerxes acela este tocmai Ahasveros, împaratul persan pomenit în cartea Esterei. Din cele scrise de istoricul grec Herodot, se pare ca ospatul din primul capitol al cartii Estera a fost prilejuit tocmai de consiliul pentru pregatirea expeditiei amintite. Patru ani mai tîrziu, cînd Xerxes mai cauta înca mîngaiere dupa dureroasa înfrîngere suferita, evreica Estera a fost facuta împarateasa (Estera 2:16).\r\nDoua femei tinere îsi dau mîna peste veacuri în dragostea lor pentru poporul evreu: Rut si Estera. Istoria Esterei este frumoasa cum numai o istorie orientala poate sa fie. Întreaga actiune se desfasoara în jurul a trei ospete: ospatul dat de Ahasveros în cinstea tuturor domnitorilor si slujitorilor lui (cap. 1 si 2), ospatul dat de Estera (cap. 7) si ospatul prilejuit de sarbatoarea Purim (cap.9). Subiectul întregii actiuni este clar: Dumnezeu stie sa-si pastreze poporul Sau în mijlocul celor mai adverse conditii. Departe de casa, pierduti în mijlocul unui urias imperiu, fara privilegii, dar cu demnitatea celor ce nu se pleaca decît în fata lui Iehova, evreii au parut o prada usoara în ochii verosului Haman, unul dintre cei mai înalti demnitari ai împaratului. Pentru potolirea propriului sau orgoliu, acest Haman pune la cale o stratagema prin care urmareste distrugerea tuturor evreilor din imperiul persan. Evenimentele se succed cu repeziciune, situatiile se schimba pe neasteptate si Haman se trezeste condamnat la moarte, iar evreii de pretutindeni sfîrsesc prin a avea doua zile de bucurie si de razbunare asupra tuturor dusmanilor lor.\r\n\r\nCuvinte cheie si terne caracteristice: Ceea ce izbeste de la început în textul cartii Estera este totala absenta a numelui lui Dumnezeu. Motivul acestei omiteri a fost cautat de-a lungul secolelor de multi comentatori. Multi spun ca traditia a hotarît ca aceasta carte sa fie citita în cursul serbarilor din zilele de Purim, care sunt „zile de ospat si de sarbatoare\" (Estera 9:21-22) si din respect pentru Dumnezeu a fost mai bine ca numele Lui sa nu fie amestecat cu bautura si veselia fara frîu.\r\n\r\nUn verset care a trecut limitele acestei carti, trecînd în patrimoniul universal este Estera 4:14: „Caci daca vei tacea acum, ajutorul si izbavirea va veni din alta parte pentru iudei, dar tu si casa tatalui tai veti pieri. ªi cine stie daca nu pentru o vreme ca aceasta ai ajuns la împaratie\".\r\nCartea ne vorbeste clar despre „providenta\" divina manifestata prin felul în care Dumnezeu îsi pastreaza mereu poporul. El lucreaza prin ordinea naturala a evenimentelor, fara sa prejudicieze vointa libera a vreunui om si fara sa întrerupa curgerea fireasca a întîmplarilor. Totusi, în spatele celor ce se întîmpla este întelepciunea Lui si grija Lui nemarginita pentru aceia pe care îi iubeste.\r\n\r\nEstera este un portret de frumusete trupeasca si sufleteasca feminina: placuta la chip si modesta la suflet (2:15), înteleapta (2:9-17; 5:1-3), ascultatoare (2:10), smerita (4:16), curajoasa (7:6), loiala si perseverenta (2:22; 8:1-2; 7:3-4).\r\n\r\nMardoheu este un exemplu de urmat pentru toti barbatii: „caci a cautat binele poporului sau si a vorbit pentru fericirea întregului sau neam\" (Estera 10:3).\r\nI. PRELIMINARII l - 5\r\na. Destituirea împaratesei Vasti, 1:1-22\r\nb. Descoperirea lui Estera, 2:1-20\r\nc. Devotamentul lui Mardoheu, 2:21-23\r\nd. Decretul lui Haman, 3:1-15\r\n\r\nII. MOMENTUL DE CRIZa, 4 - 5\r\na. Mardoheu apeleaza la Estera, 4:1-14\r\nb. Raspunsul Esterei, 4:15-17\r\nc. Bunavointa împaratului, 5:1-8\r\nd. Aroganta criminala a lui Haman, 59-14\r\n\r\nIII. IZBaVIREA PROVIDENtIALa 6-10\r\n\r\nÎNFRÎNGEREA LUI HAMAN\r\na. Haman umilit, 6:1-7:10\r\nb. Haman este spînzurat, 7:1-10\r\nc. Decretul lui Ahasveros si Mardoheu, 8:1-17\r\nd. Razbunarea asupra dusmanilor, 9:1-19\r\ne. Instituirea sarbatorii Purim, 9:20-32\r\nf. Faima si cinstea lui Mardoheu, 10:1-3" },
    //IOV
    { "IOV", "Explicatia Cartii: Iov\r\n\r\nÎn toate traducerile Bibliei cartea aceasta poarta numele personajului ei central: Iov. A fost acest Iov un personaj real? Da, profetul Ezechiel îl considera drept o personalitate proeminenta demna sa fie pusa alaturi de Noe si Daniel (Ezec. 14:14, 20), iar apostolul Iacov îl citeaza în Noul Testament ca pe un exemplu de rabdare în suferinta (Iacov 5:11).\r\nSe pare ca aceasta carte a fost scrisa initial în proza de Elihu, unul dintre personajele cartii si ca a fost tradusa în ebraica si versificata mai tîrziu de Moise. Elihu prefera sa ramîna anonim ca autor, tot asa cum a preferat sa nu se prezinte prea mult pe sine nici în textul cartii.\r\nIov a trait înainte de vremea lui Avraam. Lungimea vietii lui este caracteristica acelei perioade. Numai dupa suferinta sa, Iov a mai trait înca 140 de ani! (Iov 42:16). În textul cartii se aminteste de o moneda numita în evrreieste: „Chesita\". Despre aceasta moneda nu mai aflam decît în textele care amintesc de viata patriarhilor (Gen. 33:19). Viata lui Iov ar trebui plasata între capitolele 11 si 12 ale cartii Geneza. tara „Ut\" în care a trait Iov poarta numele unuia dintre nepotii lui Noe (Gen. 22:20-21).\r\nCartea lui Iov nu este numai o carte de istorie. Ea este o poema dramatica care ni-l prezinta pe Iov în aspectul sau filosofic. Tema cartii este realitatea si sensul suferintei în viata oamenilor lui Dumnezeu. Scena actiunii cuprinde cerul si pamîntul deopotriva. Dumnezeu si Satan se înfrunta în lumea nevazuta, iar consecintele se fac resimtite în lumea noastra prin încercari, binecuvîntari si suferinte. Cartea lui Iov este o perla a literaturii universale. Profunzimea ideilor ei transcend timpul si civilizatiile, fiind la fel de actuale astazi ca si în vremea în care au fost asezate pe hîrtie.\r\nCartea ni-l prezinta pe Iov în cîteva ipostaze distincte: Iov - omul neprihanit persecutat de Diavol din pricina neprihanirii lui (Iov 1:1-2:10), Iov - omul plin de sine care se cearta neîncetat cu prietenii sai si cu Dumnezeu (Iov 2:11-31:40), Iov - omul care se pocaieste în fata maretiei lui Dumnezeu (Iov 32: l - 42:6), Iov - omul pus în slujba pentru recuperarea celorlalti (Iov 42:7-9) si Iov - omul binecuvîntat din nou de Domnul (Iov 42:10-17).\r\n\r\nCartea prezinta doua motive pentru suferinta lui Iov: pe de o parte, avem de a face cu suferinta ca si consecinta a înfruntarii nevazute dintre Dumnezeu si Satan, iar pe de alta parte, avem de a face cu suferinta ca o disciplinare a mîndriei si a unei prea mari încrederi în sine. Daca toata actiunea s-ar fi încheiat odata cu Iov 2:10, am fi ramas cu parerea ca Iov a fost un om desavîrsit: „În toate acestea, Iov n-a pacatuit de loc cu buzele lui\". Actiunea cartii continua însa si noi descoperim treptat ca Iov este un om remarcabil printre semeni, dar lipsit de perfectiune cînd este confruntat cu standardul lui Dumnezeu. Cînd Dumnezeu intervine, dupa interminabilele înfruntari dintre Iov si cei trei prieteni ai sai, glasul Sau rasuna din mijlocul furtunii plin de mînie acuzatoare: „Cine este cel ce îmi întuneca planurile, prin cuvîntari fara pricepere?\" (Iov 38:2). Fata în fata cu Dumnezeu Creatorul, Iov se pleaca într-o pocainta sincera si adînca: „Da, am vorbit, fara sa înteleg, de minuni, care sînt mai presus de mine si pe care nu le pricep\". „Urechea mea auzise vorbindu-se despre Tine, dar acum ochiul meu Te-a vazut. De aceea mi-e scîrba de mine si ma pocaiesc în tarîna si în cenusa\" (Iov 42:1-6). În fata lui Dumnezeu, nimeni nu poate sa se laude: „Caci toti au pacatuit si sînt lipsiti de slava lui Dumnezeu\" (Romani 3:23). Chiar si cel mai bun dintre oameni merita pedeapsa divina. Chiar si cel mai stralucit exemplar uman are nevoie de har si de iertare: „Dar stiu ca Rascumparatorul meu este viu, si ca se va ridica la urma pe pamînt. Îl voi vedea si-mi va fi binevoitor\" (Iov 19:25-27).\r\n\r\nCartea Iov este unul dintre mesajele lui Dumnezeu adresate omenirii. Ea ne spune ca suferinta este de multe ori un sol pe care ni-l trimite El pentru a ne ajuta sa ne vedem corect pe noi însine si sa ne dam seama ca sîntem neputinciosi si avem nevoie de mîntuire: „Dumnezeu vorbeste însa, cînd într-un fel, cînd într-altul, dar omul nu ia seama. ªi prin durere este mustrat omul în culcusul lui, cînd o lupta necurmata îi framînta oasele\" (Iov 33:13-19), „Dar Dumnezeu scapa pe cel nenorocit prin nenorocirea lui, si prin suferinta îl înstiinteaza\" (Iov 35:15).\r\n\r\nCîta vreme cauta sa se dovedeasca nevinovat înaintea lui Dumnezeu, Iov se umple de vinovatie: „Este vreun om ca Iov, care sa bea batjocura ca apa\" (Iov 34:7). Cît timp se lauda în fata lui Dumnezeu, Iov este vinovat de pacatul mîndriei, iar pentru cei mîndrii cerul se închide: „Sa tot strige ei atunci, caci Dumnezeu nu raspunde, din pricina mîndriei celor rai. Degeaba striga, caci Dumnezeu n-asculta, Cel Atotputernic nu ia aminte\" (Iov 35:12-13).\r\nI. PROLOG l - 2\r\na. Iov - neprihanit în belsug, 1:1-5\r\nb. Satan - rautacios si pornit pe rau, 1:6-19\r\nc. Iov - neprihanit în nenorocire, 1:20-22\r\nd. Satan - perseverent în rautate, 2:1 -8\r\ne. Iov - neprihanit pîna la capat, 2:9-13\r\n\r\nII. DIALOG 3 - 42\r\na. Plîngerea lui Iov, 3\r\nb. Prima triada de dialoguri, 4-14\r\nc. A doua triada de dialoguri, 15 - 21\r\nd. A treia triada de dialoguri, 22-31\r\ne. Elihu intervine, 32 - 37\r\nf. Dumnezeu intervine, 38-41\r\n\r\nIII. EPILOG 42:7-17\r\na. Iov - laudat de Dumnezeu, 42:7\r\nb. Cei trei prieteni condamnati de Domnul, 42:8\r\nc. Iov - suferinta lui înceteaza, 42:10\r\nd. Iov - rasplatit cu binecuvîntare, 42:11-17" },
    //PASALMI
    { "PSALMI", "Explicatia Cartii: Psalmi\r\n\r\nPrimele fragmente de traducere a Bibliei în limba româna au fost din cartea psalmilor. Nicolae Iorga spunea ca: „Neamul românesc s-a nascut în tinda Bisericii\". „Psaltirea\" a fost abecedarul pe care neamul nostru a învatat sa citeasca. De ce s-a început traducerea Bibliei în limba româna cu psalmii? Raspunsul nu este greu de dat. Cartea aceasta este cea mai mare din Biblie si cea mai universal citita si studiata. Ea cuprinde cel mai larg evantai de trairi sentimentale posibile.\r\n\r\nLuati Biblia în mîna si deschideti-o la mijloc. Veti da negresit de cartea Psalmilor. Este normal sa fie asa deoarece psalmii sînt inima Bibliei. Aici sînt adunate toate trairile sufletesti ale copiilor lui Dumnezeu de-a lungul secolelor.\r\nNumele evreiesc este „Tehilim\" - „laudele\". Numele din limba româna vine de la traducerea în limba greaca. „Psalmoi\" s-ar putea traduce prin „cîntari acompaniate cu instrument muzical\".\r\nNici una dintre cartile Bibliei nu are mai multi autori decît cartea Psalmilor. În general se spune ca psalmii îi apartin lui David. El i-a scris pe majoritatea din ei si tot el da nota generala a cartii. Dar asta nu înseamna ca n-au existat si alti autori ai psalmilor. Daca este adevarat ca David a scris 73 din numarul total de 150, atunci tot atît de adevarat este ca Asaf, conducatorul cîntaretilor lui David, a scris 12, Solomon a scris 2, Moise a scris l (Ps.90), iar alti 10 sînt productia fiilor lui Core (Num. 26:9-11). Nu va pierdeti prea multa vreme însa cu identificarea autorilor. Cititi psalmii ca pe un dar pe care vi-l face Dumnezeu însusi. Cautati sa intrati în atmosfera lor, cautati sa le simtiti trepidatia adînca si identificati-i pe aceia care va vor prinde bine la vreme de nevoie.\r\nPsalmii acopere o perioada care debuteaza cu viata lui Moise si se întinde pîna dupa revenirea din robia babiloniana, pe vremea lui Ezra si Neemia. Majoritatea psalmilor însa sînt produsi de David în timpul vietii lui de umblare cu Domnul. Omul acesta care a fost cioban la oi, cîntaret vestit în Israel, trubadur la curtea împaratului, ginere al împaratului, apoi fugar haituit si aflat mereu în primejdie pîna cînd Dumnezeu l-a asezat chiar pe el însusi pe tronul Israelului, a strabatut atâtea suisuri si coborîsuri ale spiritului încît a devenit instrumentul la care a cîntat însusi Dumnezeu. Psalmii lui David sînt de fapt cîntari divine inspirate din cer pentru a înfrumuseta viata plina de necazuri a pamîntului. Din acest punct de vedere, psalmii sînt atemporali, ei vin la noi din rezonantele timpului, ne învaluie cuceritor si vor ramîne cu noi vesnic.\r\nMetoda de interpretare a psalmilor: Poezia nu este o arta precisa, ca matematica. Dincolo de faldurile ei armonioase patrundem adesea în straturi ascunse ale realitatii, pe care nu le întrezarim în viata decît în situatii extreme. Cum sa fim siguri ca ceea ce simtim atunci cînd citim psalmii este ceea ce este acolo si nu ne amagim lasîndu-ne purtati de valurile fanteziei?\r\n\r\nExista cîteva principii pe care trebuie sa le tinem minte:\r\n\r\n(1) Cînd exista un titlu al psalmului care îl leaga de un anumit eveniment istoric, psalmul trebuie citit, înteles si interpretat în lumina acelei circumstante particulare. De exemplu, 13 psalmi sînt localizati în diverse perioade ale vietii lui David: Ps. 59, 56, 34, 142, 52, 54, 57, 60, 51, 3, 63, 7, 18. Ordinea aceasta nu pare normala, dar ea urmareste cronologic evenimentele vietii lui David asa cum sînt ele redate de cartea 1 Samuel.\r\n\r\n(2) Unii dintre psalmi sînt asociati cu anumite aspecte din închinaciunea poporului Israel (vezi Ps. 5:7; 66:13; 68:24, 25) si trebuiesc întelesi în ansamblul procesiunilor liturghice de la Templu.\r\n\r\n(3) Multi dintre psalmi sînt de fapt inspiratii profetice care anticipeaza venirea lui Cristos pe pamînt. ªi atunci însa, ei au aparut într-un anumit context istoric legat de evenimente contemporane autorului lor. Aceste evenimente trebuiesc studiate cu atentie pentru a ne ajuta sa avem o întelegere mai clara a textului.\r\n\r\nInterpretarea psalmilor este o problema a mintii, trairea mesajului lor este o problema a inimii. S-a spus despre cartea Psalmilor ca „este un rîu de mîngîiere în apa caruia s-au adunat lacrimile varsate de nenorocitii trecutului pentru a racori arsita din sufletele celor ce sufera astazi\". Sau ca este „gradina tuturor florilor cu petale parfumate, chiar daca unele dintre ele cresc numai în mijlocul spinilor\". A fost comparata cu „un desavîrsit instrument muzical din corzile caruia pot prinde fiinta si jalea si triumful, si deznadejdea si biruinta, si teama si încrederea nestramutata, si tristetea si bucuria fara margini\".\r\nLa o privire mai atenta observam ca avem de a face cu o culegere de cinci carti ale psalmilor, fiecare dintre ele terminate cu o doxologie (o proslavire a numelui lui Dumnezeu). Numele de cinci le-a adus imediat o asemuire cu Pentateucul lui Moise. Comentatorii numesc cele cinci carti ale psalmilor dupa numele corespunzator al uneia dintre cele cinci carti scrise de el.\r\n\r\na. Sînt psalmi ai genezei (1-41) al carui subiect predominant este omul,\r\nb. psalmi ai exodului, cuprinzînd mai ales cîntari de eliberare (42-72),\r\nc. psalmi levitici, asociati cu slujba de la Templu (73-89),\r\nd. psalmi asociati cu peregrinarile din cartea Numeri (90-106) si\r\ne. psalmi deuteronomici sau ai proslavirii dragostei divine (107-150).\r\n\r\nUn alt fel de a clasifica psalmii este dupa specificul mesajului lor. Exista astfel:\r\n\r\n(1) psalmi de învatatura (1, 19, 39),\r\n(2) psalmi de lauda (8, 29, 93, 100),\r\n(3) psalmi de multumire (30, 65, 103, 107, 116),\r\n(4) psalmi de pocainta (6, 32, 38, 51, 102, 130, 143),\r\n(5) psalmi prin care strabate încrederea (3, 27, 31, 46, 56, 62, 86),\r\n(6) psalmi ai necazurilor (4, 13, 55, 64, 88),\r\n(7) psalmi ai dorintelor (42, 63, 80, 84, 137),\r\n(8) psalmi care recapituleaza istoria (78, 105, 106) si\r\n(9) psalmi profetici (2, 16, 22, 24, 40, 45, 68, 69, 72, 97, 110, 118)\r\n\r\nContinutul psalmilor este pus adesea în forme de o desavîrsita frumusete. Iata, de exemplu, psalmul 119 al carui autor se presupune ca a fost carturarul Ezra. Pentru cititorul neavizat este greu sa nu te ratacesti în multimea celor 176 de versete ale lui. În originalul ebraic însa, psalmul este o succesiune de grupe de cîte 8 versete, fiecare începînd cu cîte una din literele alfabetului evreiesc. Avem de a face deci cu un acrostih alfabetic alcatuit dintr-o serie de octete. Analiza continutului acestui psalm devine dintr-o data mai usoara si mai fructuoasa.\r\nPsalmii sînt cartea tuturor oamenilor, tuturor sentimentelor si a tuturor situatiilor. Oricine, oriunde si în orice stare s-ar gasi se poate identifica cu ceea ce se gaseste în textul psalmilor. Cel sarac si parasit, cel bolnav si aflat în suferinta, cel exilat departe de cei dragi si de tara, cei aflati mereu în pericol, cei pacatosi se pot regasi în oglinda psalmilor. Este loc însa acolo si pentru cei iertati, pentru cei biruitori, pentru cei prabusiti în adorare, pentru cei atasati total de Domnul si locasurile Sale, pentru cei pierduti lor însile si daruiti pe vecie Creatorului.\r\n\r\nCartea Psalmilor este si cartea Legii desavîrsite a lui Dumnezeu. Ea ne arata cum putem sa ne delectam în ea si sa o facem „candela pentru picioarele noastre, lumina pe carare\" si desfatare „mai dulce ca fagurul de miere pentru cerul gurii\".\r\n\r\nCa si celelalte carti ale Bibliei, cartea Psalmilor ne vorbeste si ea despre planul pe care l-a facut Dumnezeu pentru mîntuirea omenirii. Dupa înviere, Domnul Isus li s-a aratat ucenicilor: „le-a deschis mintea ca sa înteleaga Scripturile\" si le-a aratat ceea ce fusese scris cu privire la El si lucrarea Sa „în Legea lui Moise, în Prooroci si în Psalmi\". Putem si noi sa refacem acest studiu exegetic facut de Domnul Isus cu ucenicii. Iata un rezumat al lui:\r\n\r\na. Despre esenta lucrarii Lui mesianice gasim scris în Ps. 22:22\r\nb. Despre slujirea Sa ca Mare Preot gasim scris în Ps. 40:6, 8; 22; 49; 110.\r\nc. Despre demnitatea Lui de împarat gasim scris în Ps. 2; 21; 45; 72.\r\nd. Despre suferintele Lui teribile gasim scris în Ps. 22 si 69\r\ne. Despre învierea Lui gasim scris în Ps. 16\r\n\r\nDintre toate cartile Vechiului Testament, cartea Psalmilor este cel mai mult citata în cartile Noului Testament. Unele dintre portiunile ei sînt parca parte integranta din Evanghelii. Iata de exemplu succesiunea celor trei psalmi: 22, 23 si 24. Cel dintîi ne prezinta suferintele pastorului cel bun care-si da viata pentru oile Sale. Vedem ridicata crucea îi auzim strigatele sfîsietoare ale Celui ce moare o moarte ispasitoare. Psalmul 23 este cunoscut de toata crestinatatea drept „psalmul pastorului\" („Domnul este Pastorul meu, nu voi duce lipsa de nimic. El ma paste la pasuni verzi si ma duce la ape de odihna, etc...”) Gasim în acest psalm oferta Domnului Isus de a conduce si de a îngriji sufletele care i se încredinteaza. Psalmul 24 ni-L prezinta pe pastor ajuns în slava gata sa rasplateasca oile care L-au urmat pîna la capat.\r\n\r\nSoarta poporului Israel este si ea descrisa într-o succesiune de psalmi. Iat-o:\r\n\r\na. Ruina poporului este descrisa în Ps. 42-49\r\nb. Rascumparatorul poporului este prezentat în Ps. 50-60\r\nc. Rascumpararea poporului este aratata în Ps. 61-72\r\n\r\nEste greu sa aratam mesajul psalmilor într-o introducere scurta ca aceasta. Va lasam dumneavoastra placerea ca dupa ce ati parcurs cu noi aceasta vedere panoramica a lor, sa va apropiati pe rînd de fiecare psalm în parte, sa-i descoperiti frumusetea si sa-i sorbiti cu nesat aroma specifica." },
    //PROVERBE
    { "PROVERBE", "Explicatia Cartii: Proverbe\r\n\r\nCartea aceasta nu trebuie judecata dupa volum, ci dupa imensitatea întelepciunii pe care o gazduieste. Mii de volume din bibliotecile lumii nu ne pot ajuta, toate la un loc, cît ne poate ajuta aceasta singura carte. Autorul ei nevazut este Dumnezeu, iar priceperea ei, desi nascuta în sferele cerului este destinata sa revolutioneze viata terestra.\r\nUn proverb este un discurs de întelepciune redus la o singura fraza.\r\nTextul cartii îl mentioneaza pe Solomon drept autor la începutul fiecareia din cele trei sectiuni ale ei (1:1; 10:1; 25:1). Despre Solomon stim deja ca a fost autorul a 3.000 de proverbe si a 1.005 cîntari (1 Împ. 4:32). Nici un om din Israel nu a fost mai potrivit sa editeze o carte de întelepciune ca acest Solomon. El s-a rugat Domnului pentru întelepciune (1 Împ. 3:5-9) si a primit-o asa cum nu i-a mai fost data nici unui om de pe fata pamîntului (1 Împ. 4:29-31). Stralucirea gîndirii lui l-a facut celebru în lumea de atunci si a atras admiratia celor veniti de la mari departari ca sa-l vada si sa-l asculte (1 Împ. 4:34; 10:1-l3, 24). Contributia lui Solomon la ridicarea Israelului a fost imensa. Capacitatea lui de sinteza, intuitia si clarviziunea lui sînt si astazi proverbiale. Asta nu înseamna ca tot ceea ce a scris Solomon a fost produsul sau nemijlocit. El însusi ne spune ca i-a placut sa zaboveasca îndelung asupra întelepciunii raspîndite de scrierile altora: „Pe lînga ca Ecleziastul a fost întelept, el a mai învatat si stiinta pe popor, a cercetat, a adîncit si a întocmit un mare numar de zicatori. Eclesiastul a cautat sa afle cuvinte placute, si sa scrie întocmai cuvintele adevarului\" (Ecles. 12:9-10). Prin aceasta, el a recunoscut ca întelepciunea nu este monopolul unui singur om, ci este darul facut de Dumnezeu oamenilor: „Cuvintele înteleptilor sînt ca niste bolduri; si, strînse la un loc, sînt ca niste cuie batute, date de un singur stapîn\" (Ecles. 12:11).\r\nAceasta culegere de proverbe a fost alcatuita prin preajma anului 931 î.Cr. de catre Solomon. Capitolele 25-29 ale cartii au fost culese mai tîrziu de Ezechia si adaugate cartii lui Solomon.\r\nÎn aceasta carte avem de a face cu mult mai mult decît cu o culegere de proverbe. De fapt sînt o varietate întreaga de procedee stilistice sub care ne-a fost transmisa „întelepciunea\" acumulata de acest împarat cu o desavîrsita capacitate artistica. Cuprinsul cartii semnaleaza diferitele modalitati de exprimare, asa ca nu le vom mai enumera aici.\r\nÎn capitolul 8 al cartii, „întelepciunea\" este personificata si descrisa în toata perfectiunea ei. Ea este de origine divina (8:22-31), este izvorul vietii biologice si spirituale (8:35, 36; 3:18), este neprihanita si adevarata (8:8, 9) si se ofera tuturor celor ce o cauta (8: 1-6, 32-35). Peste veacuri aceasta „întelepciune\" s-a întrupat în persoana Domnului Isus Cristos, „în care sînt ascunse toate comorile întelepciunii si ale stiintei\" (Col. 2:3). Cartea proverbelor este o anticipare a întîlnirii cu Cristos, „care a fost facut de Dumnezeu pentru noi întelepciune, neprihanire, sfintire si rascumparare\" (1 Cor. 1:30; conform lui 1 Cor. 1:22-24). Iata schita acestei carti:\r\nI. LAUDA ÎNTELEPCIUNII l - 9\r\n15 sonete. Introducere (1:1-9);\r\n\r\nAdemenire din partea pacatosilor (1:10-19)\r\nÎntelepciunea care elibereaza (2:1-22)\r\nFolosul temerii de Dumnezeu (3:1-10)\r\nÎntelepciunea, rasplata suprema (3:11-20)\r\nÎntelepciunea, suprema siguranta (3:21-26)\r\nÎntelepciunea si viclenia (3:27:35);\r\nÎntelepciunea, suprema mostenire (4:1-9)\r\nCele doua cai (4:10-19)\r\nÎntelepciunea si sanatatea (4:20-27)\r\nFemeia straina (5:1-23)\r\nDespre chezasie (6:1-11)\r\nLenesul (6:6-11)\r\nCel ce seamana certuri (6:12-19)\r\nFereste-te depreacurvie (6:20-35)\r\nCasa întelepciunii si Casa Nebuniei (cap. 9)\r\n\r\n2 monologuri\r\n\r\nAvertismentul întelepciunii (1:20-33)\r\nÎntelepciunea si femeia straina (cap. 7 si 8)\r\n\r\nII. MAXIMELE ÎNtELEPCIUNII 10 - 24\r\n375 de proverbe sau aforisme în forma de afirmatii care se contrasteaza, se completeaza sau se compara reciproc (10:1-22:16)\r\n\r\n16 epigrame. Introducere (22:17-21);\r\nEpigrame amestecate (22:22-29);\r\nPericolul lacomiei (23:1-3);\r\nDesertaciunea bogatiilor (23:4-5)\r\nGazda vicleana (23:6-8);\r\nEpigrame amestecate (23:9-18);\r\nÎmbuibarea vinovata (23:19-21);\r\nTrei ziceri (23:22-25)\r\nCursa celei stricate (23:26-28);\r\nVin si Vai (23:29-35);\r\nEpigrame amestecate (24:1-10);\r\nScapa-l daca poti (24:11-12);\r\nÎntelepciunea si mierea (24:13-14);\r\nPatru epigrame (24:15-22);\r\nImpartialitate (24:23-25);\r\nTrei ziceri (24:26-29);\r\nOgorul lenesului (24:30-34)\r\n\r\nIII. ALTE MAXIME 25-31\r\n7 epigrame si proverbe înmanunchiate.\r\nÎmparatul (25:1-7);\r\nDiferite (25:6-26:2);\r\nDespre nebuni 26:3-12);\r\nLenesul (26:13-16);\r\nDezbinatorii (26:17-26);\r\nDiverse (26:27-27:22)\r\nGospodarul bun (27:23-27)\r\n55 de proverbe sau aforisme În forma de perechi care se contrasteaza, se completeaza sau se compara reciproc (cap.28 si 29)\r\nCele treisprezece ziceri ale lui Agur (cap.30)\r\nSpusele mamei lui Lemuel (cap.31:1-9)\r\nUn acrostih despre femeia vrednica de cinste (cap.31:10-31)" },
    //ECLESIASTU
    { "ECLESIASTUL", "Explicatia Cartii: Eclesiastul\r\n\r\nCartea Eclesiastul este o confesiune a unui suflet care a esuat în cautarea lui dupa fericire. Textul este plin de pesimism si dezamagire. Excluzîndu-l pe Dumnezeu din viata lui, traind la nivelul „de sub soare\", si netînînd cont de întreaga panorama a vietii, a carei explicatie este coerenta numai daca iei în consideratie dimensiunea ei eterna, cautatorul din cartea Eclesiastul a ajuns la capatul drumului „cu mîna goala\". Daca vreti sa stiti unde poate ajunge cel ce le are pe toate; bogatie, lux, femei, bautura, cîntece si veselie, atunci cititi neaparat aceasta carte. Ea este una dintre cele mai neîntelese si mai nedreptatite carti din cartile Bibliei. Cei pesimisti s-au laudat ca au gasit în ea dovada inutilitatii cautarilor umane. Scepticii s-au agatat de ea ca de o dovada în sprijinul parerii ca dupa moarte nu mai urmeaza nimic, ci doar o permanenta nefiintare. Altii au citat-o atunci cînd au vrut sa demonstreze ca dupa moarte sufletul „doarme\" si nu este constient de nimic în intervalul scurs între clipa mortii si ceasul învierii. În afara de acesti entuziasti direct interesati, nu se gasesc, mai ales printre crestini prea multi oameni care sa o citeasca fara sa se mire ce cauta ea în Biblie. Pentru multi oameni sinceri, textul Eclesiastului, atunci cînd nu contravine direct cu învatatura Noului Testament, este cel putin bizar si greu de armonizat cu „întregul\" Scripturii. Nadajduim ca rîndurile care urmeaza îi vor face pe multi sa-si schimbe parerea pe care o au despre continutul ei.\r\n\r\nAceasta extraordinara carte de filosofic poate fi asemuita cu marile lucrari muzicale ale lui Johan Sebastian Bach: ele sînt pline de prabusiri si zvîrcoliri minore, dar se încheie întotdeauna cu înaltatoarele tonalitati ale gamelor majore.\r\nÎn original cartea s-a numit „Kohelet\". Acesta este un termen rar al limbii ebraice folosit în Biblie numai în cartea Eclesiastul (1:1, 2, 12; 7:27; 12:8-10). Termenul este un derivat de la „kahal\" - a convoca o adunare, a aduna împreuna. Sensul titlului este deci: „cel ce se adreseaza unei adunari, predicatorul\". Traducerea greaca Septuaginta foloseste titlul: Ecclesiaites, de unde s-a derivat numirea pentru Biserica: „eclesia\". Numirea româneasca ar fi trebuit deci sa fie: „Cel ce vorbeste adunarii\", dar traducatorul a preferat sa translitereze titlul din limba greaca.\r\nFara îndoiala ca autorul a fost împaratul Solomon. Cel ce scrie cartea se intituleaza pe sine: „Fiul lui David, împaratul Ierusalimului\" (Ecles. 1:1). Incursiunile autorului în placerile de tot felul (2:1-3), în realizari marete (2:4-6) si într-o viata de bogatie fara asemanare (2:7-10) îl identifica fara echivoc pe Solomon.\r\nSolomon a scris aceasta carte la batrinete, spre sfîrsitul vietii lui, probabil în preajma anului 935 î.Cr.\r\nTraditia evreiasca spune ca Solomon a scris cartea Cîntarea Cîntarilor în tinerete, cartea Proverbelor la maturitate si cartea Eclesiastul spre apusul vietii, cînd a ajuns sa fie coplesit de regrete pentru anii irositi în placerile carnii si în idolatrie (1 Împ. 11).\r\nCa sa stabilim de la început caracterul acestei carti trebuie sa spunem ca ea nu este o carte „inspirata\" de Dumnezeu, dar ca autorul ei a fost „inspirat\" sa o scrie. Nu trebuie sa mergem în afara Bibliei ca sa aflam cum gîndeste lumea. Dumnezeu a vrut ca sa ne întîlnim cu gîndirea celor fara Dumnezeu chiar pe paginile Scripturii. Mai mult decît atît, El a vrut sa fim în stare sa ne confruntam cu ea. Eclesiastul ne aseaza înainte tot ceea ce mintea umana a reusit sa afle în cautarile ei dupa fericire si dupa sensul real al existentei. Argumentele folosite în aceasta carte nu sînt argumentele lui Dumnezeu, ci argumentele mintii umane. Aceasta explica pe deplin aparitia unor pasaje ca acelea din Ecl. 1:15; 2:24; 3:3, 4, 8, 11, 19, 20; 8:15 care sînt într-un contrast flagrant cu tot restul scrierilor sfinte. Solomon îsi descrie incursiunile vietii lui în cautarea dupa sens si valoare. Pentru început, el nu neaga, ci neglijeaza pur si simplu dimensiunea transcendentala, si-si directioneaza cautarile înspre lumea stiintei (Ecl. 1:4-11), înspre lumea filosofiei (1:12-18), înspre lumea senzuala a placerilor: petreceri (Ecl. 2:1), bautura (Ecl. 2:3), realizari materiale marete (2:4-7), bogatie si muzica (2:8), femei (2:8). El încearca rînd pe rînd: materialismul (Ecl. 2:12-26), fatalismul (Ecl. 3:1-15), deismul (Ecl. 3:1-4), religia naturii (Ecl. 5:1-8), avaritia (5:9-6:12) si chiar moralitatea (7:1-11:8). Toate acestea se dovedesc însa în final: „Desertaciune a desertaciunilor, o desertaciune a desertaciunilor! Totul este desertaciune si goana dupa vînt!\" (Ecl. l:2, 14).\r\n\r\nCuvinte cheie si teme caracteristice: Cuvîntul care caracterizeaza întreaga carte este „desertaciune\". Scopul nedeclarat al cartii este acela de a demonstra zadarnicia cautarii dupa fericire în domeniul existentei vremelnice. Fericirea nu poate fi gasita daca este cautata ca un scop în sine, pentru ca ea nici nu exista ca ceva de sine statator. Fericirea este o stare care însoteste o anumita atitudine. Iar atitudinea aceasta este împlinirea vointei vesnice a lui Dumnezeu. Este important ca Eclasiastul leaga apelul sau pentru considerarea responsabilitatilor fata de Dumnezeu de vîrsta tineretii (!). Cine vrea sa traiasca fericit este bine sa nu repete cautarile inutile ale Eclesiastului si sa înceapa din tinerete sa intre în „dimensiunea\" ascultarii de Dumnezeu.\r\n\r\nCele mai cunoscute si mai citate versete sînt:\r\n\r\n„O desertaciune a desertaciunilor. Totul este desertaciune!\" (Ecl.l:2).\r\n„Toate îsi au vremea lor si fiecare lucru de sub ceruri îsi are ceasul lui\" (Ecl. 3:1).\r\n„Funia împletita în trei nu se rupe usor\" (Ecl. 4:12b).\r\n„Arunca-ti pîinea pe ape si dupa multa vreme o vei gasi iarasi!\" (Ecl. 11:1).\r\n„Dar adu-ti aminte de Facatorul tau în zilele tineretii tale, pîna nu vin zilele cele rele si pîna nu se apropie anii, cînd vei zice: „Nu gasesc nici o placere în ei\" (Ecl. 12:1).\r\n„Sa ascultam dar încheierea tuturor învataturilor: Teme-te de Dumnezeu si pazeste poruncile Lui. Aceasta este datoria oricarui om\" (Ecl. 12:13).\r\nAcest mesaj ar putea fi rezumat în trei afirmatii:\r\n\r\n(1) Cînd observi viata umana cu ciclurile ei aparent lipsite de semnificatie (Ecl.l:4-11) si cu paradoxurile ei inexplicabile (4:1; 7:15; 8:8), ajungi neaparat la convingerea ca totul este lipsit de sens si de semnificatie, deoarece este imposibil sa gasesti în toate macar un singur scop în jurul caruia sa-ti organizezi existenta.\r\n\r\n(2) Cu toate acestea, îti dai seama ca este bine ca viata sa fie traita din plin, caci este un dar pe care ti l-a facut Dumnezeu (3:12-13; 3:22; 5:18-19; 8:15; 9:7-9), numai ca trebuie sa traiesti cu grija stiind ca...\r\n\r\n(3) vei da socoteala Facatorului tau în ziua în care te va chema la judecata (3:16-17; 12:14).\r\n\r\nNu va lasati pacaliti de cei care spun ca n-avem nevoie de Eclesiastul în Biblie si ca standardul spiritual al cartii este mult sub nivelul Noului Testament. Lumea de azi este plina de „înfumurati super-spirituali\" cu iz de superioritate care cad mereu în aceleasi capcane, repeta mereu aceleasi cautari iluzorii si gusta mereu aceleasi dezamagiri amare ale pacatului. Toate acestea din cauza ca ei n-au citit niciodata o carte atît de scandalos de sincera ca aceasta si pentru ca ei n-au avut ocazia sa asculte un predicator atît de revoltator de onest ca acest Kohelet din vechime.\r\nIntroducere (1:1)\r\n\r\nI. Prima predica 1-2\r\na. Prezentarea tezei de baza: zadarnicia eforturilor si înfaptuirilor umane (1:2-3)\r\n\r\nb. Demonstrarea tezei din exemplul experientei proprii\r\nLipsa desens a ciclurilor vietii si istoriei (1:4-11)\r\nLipsa de importanta a întelepciunii si filosofiei omenesti (1:12-18)\r\nLipsa de satisfactie din avere si viata de placeri (2:1 -11)\r\nLipsa de viitor în perspectiva unei morti sigure si necrutatoare (2:12-17)\r\nLipsa de sens a ostenelii în munca (2:18-23)\r\n\r\nc. Concluzia: Multumire cu ceea ce avem în clipa de fata (2:24-26)\r\n\r\nII. A doua predica 3-5\r\na. Prezentarea tezei de baza: Problema vietii si a mortii (3:1 -22)\r\nb. Neîmplinirile si dezamagirile vietii pe pamînt (4:1-16)\r\nc. Concluzia: Zadarnicia goanei dupa fericire (5:1-20)\r\nDesertaciunea ritualului religios (5:1-7)\r\nRealitatea iluzorie a bogatiei (5:8-17)\r\nFericirea unui suflet multumit (5:18-20)\r\n\r\nIII. A treia predica 6 - 8\r\na. Prezentarea tezei de baza: Insuficienta materialismului (6:1-12)\r\nb. Cîteva sfaturi întelepte (7:1-8:11)\r\nc. Concluzia: Dumnezeu face dreptate pentru tot (8:12-17)\r\n\r\nIV. A patra predica 9-12:8\r\na. Prezentarea tezei de baza: Siguranta mortii, nesiguranta vietii (9:1-18)\r\nb. Demonstrarea tezei (10:1-20)\r\nc. Concluzia: Viata ca o oportunitate cu responsabilitati (11:1-12:8)\r\n\r\nV. Epilog: O prezentare a temei cartii 12:9-14\r\na. Ciclurile vietii (12:9-12)\r\nb. Responsabilitatea vietii: teama de Dumnezeu si ascultarea de voia Lui (12:13-14)" },
    //CINTAREA CINTARILOR
    { "CÂNTAREA CÂNTĂRILOR", "Explicatia Cartii: Cantarea Cantarilor\r\n\r\nCîntarea Cîntarilor este o oda a iubirii care ridica dragostea dintre barbat si femeie la nivelul stabilit initial de Facatorul nostru. Adevarata religie condamna deopotriva si ascetismul si depravarea. Culmea iubirii este reciprocitatea de sentimente si trairi dintre un sot si o sotie în cadrul sfînt al familiei.\r\nªi în originalul evreiesc si în greaca si în latina, cartea poarta acelasi nume: „Shir Hashiram\", „Asma Asmaton\" si „Canticum Canticorum\" care se traduc toate în româneste prin „Cîntarea Cîntarilor\". Numirea subliniaza valoarea textului care trece drept un superlativ al artei literare evreiesti.\r\nTextul însusi ne spune ca aceasta carte este scrisa de Solomon. (1) Numele lui apare de sapte ori în continutul ei (1:1, 5; 3:7, 9, 11; 8:11, 12). (2) Conditiile de abundenta si belsug pe care numai unul ca Solomon le-a avut sînt descrise în capitolul 3:6-11. (3) Cetatile pomenite în text fac aluzie la existenta unei tari care nu traise înca tragedia divizarii.\r\n\r\nTextul din 1 Împarati 4:32, 33 ne spune ca Solomon a compus 1.005 cîntari si a avut o cunoastere deosebita despre animalele si plantele din natura. În textul Cîntarii Cîntarilor sînt pomenite 21 de feluri de plante si 15 specii de animale.\r\nCîntarea a fost compusa în tineretea lui Solomon, probabil cam prin anul 965 î.Cr.\r\nCa în toate marile povesti de dragoste din lume, faptele acestei întîmplari sînt simple si aparent, banale. Ceea ce le confera o incomensurabila valoare este bogatia de sentimente umane care le însotesc. Iata mai întîi datele acestui poem de dragoste:\r\n\r\nÎmparatul Solomon avea o vie în tinutul deluros al lui Efraim, cam la 80 de kilometri de Ierusalim (8:11). El a lasat-o în îngrijirea unor arendasi (8:11) care erau de fapt o familie constituita din mama, doi fii (1:6), si doua fete: Sulamita (6:13) si înca o sora mai mica a ei (8:8). Aceasta Sulamita era un fel de „cenusareasa\" a familiei (1:5), foarte frumoasa, dar nebagata în seama de nimeni. Fratii ei o tratau cu toata asprimea punînd-o la munci grele, asa ca ea nu si-a putut pastra si îngriji prea mult frumusetea chipului (1:6). Era pusa sa tunda via si sa aseze curse pentru vulpile mici care stricau via (2:15). Uneori era trimisa sa pasca oile (1:8). Viata sub arsita soarelui o facuse sa aiba pielea complet arsa de soare (1:5).\r\n\r\nÎntr-o zi, un chipes strain s-a apropiat de via pazita de arendasi si a stat de vorba cu Sulamita. Era împaratul Solomon, care fara sa-si faca cunoscuta adevarata identitate a stat de vorba cu Sulamita. Sub privirile lui pline de admiratie si la auzul vorbelor lui mestesugite de cuceritor de inimi, Sulamita a simtit pentru prima data necaz pentru înfatisarea ei neîngrijita (1:6). Crezîndu-l un cioban pribegind cu oile, ea l-a întrebat despre turma lui (1:7). El i-a raspuns numai în termeni generali (1:8), dar a continuat sa-i vorbeasca asemeni unui îndragostit (1:8-10) si sa-i promita daruri ca semn al dragostei si seriozitatii lui (1: 11). Pe scurt, strainul i-a cucerit inima si a plecat promitîndu-i ca va reveni nu peste multa vreme. În absenta lui, imaginatia ei s-a aprins de dor. A început sa-l viseze noaptea, nadajduind în fiecare dimineata ca el va veni împreuna cu zorile (3:1). Într-un sfîrsit, el s-a întors asa cum promisese, de data aceasta cu toata splendoarea si fala sa împarateasca si a luat-o de sotie (3:6-7).\r\n\r\nPe aceasta întîmplare simpla, Solomon brodeaza toata maiestria lui ornamentala specifica orientului. Dragostea nu este o problema de fapte, ci de trairi emotionale intense si niciodata nu a existat vreun om mai priceput ca Solomon în descrierea acestor sentimente inefabile care scapa cercetarii rationale, dar ne aprind pîna la incendiere inima. Cartea este plina de comparatii, metafore, alegorii si personificari fermecatoare.\r\nSe pare ca destinatia initiala a acestui poem a fost educatia noilor casatoriti în tainele si consolidarea dragostei. Ca si celelalte carti poetico-didactice, Cîntarea Cîntarilor a fost unul din manualele dupa care se facea educatia în Israel. Cartea a avut însa si o alta destinatie. Ea era citita în public la unele din sarbatorile anului calendaristic ebraic. În acel cadru solemn, cartea nu mai era numai un manual al dragostei, ci se ridica infinit mai sus, devenind o alegorie a iubirii dintre Dumnezeu si poporul Sau pe care si l-a ales pentru vecie. Aceasta lectura publica pare, daca nu total nepotrivita, cel putin discutabila, societatii moderne de astazi. Nu trebuie sa uitam însa ca în Israel imaginea casniciei era sfînta si ca ea fusese folosita de însusi Dumnezeu în dialogul Lui cu poporul (Osea 2:1-20; 5:7; Isaia 50:1; Ieremia 3:1-25). În Noul Testament Pavel si Ioan folosesc si ei alegoria casniciei pentru a ilustra relatia dintre Domnul Isus si Biserica (Efes. 5:25-27; Apoc. 19:7; etc.) Probabil ca si aceasta libertate a lor ar fi fost criticata de „puritanii\" pretentiosi de astazi, dar sarcina lor a fost mult usurata de existenta acestei Cîntari a Cîntarilor în canonul Vechiului Testament.\r\nExista trei teorii asupra felului în care trebuie înteleasa aceasta carte:\r\n\r\n(1). Interpretarea naturalista care nu vede în aceasta carte nimic altceva decît o colectie de cîntece erotice de dragoste, asezate laolalta din cauza subiectului si frumusetii lor comune. Aceasta teorie face din includerea cartii în canonul religios evreiesc un fapt inexplicabil si fara sens. Cînd ne gîndim cît de mult au pretuit evreii scrierile lor sfinte si cu cîta atentie si-au pastrat ei aceste „oracole\" vesnice, aceasta teorie îsi pierde foarte repede orice valabilitate.\r\n\r\n(2). Interpretarea alegorica, care transforma întreaga carte într-o succesiune de tablouri simbolice cu sensuri ascunse si sublime. Excesul de imaginatie îi face însa pe sustinatorii acestei interpretari sa nege fundamentul istoric real al cartii si-i determina sa caute în fiecare mic detaliu al textului corespondente spirituale desavîrsite. Un astfel de comentator a „vazut\" în parul bogat al femeii descrise în text nici mai mult, nici mai putin decît multimea neamurilor care au încadrat Israelul în lucrarea de mîntuire!\r\n\r\n(3). Interpretarea tipologica gaseste o cale de mijloc între cele doua interpretari de mai sus, pastrînd ce este bun în fiecare si refuzînd excesele extremismelor. Conform acestei interpretari, cartea are un fundament istoric real cu o desfasurare cronologica de evenimente. Întîmplarea este asezata în sapte „tablouri\" teatrale independente si interdependente în acelasi timp care descriu dragostea ideala dintre un barbat si o femeie. Frumusetea sfînta a acestei uniri planuite în tiparul ei de însusi Dumnezeul creatiei, este preluata într-un plan mai înalt pentru a ilustra în „tip\" intensitatea iubirii care trebuie sa existe între Dumnezeu si Israel, în interpretarea evreiasca si între Cristos si biserica, în interpretarea crestina.\r\nPersonajele „distributiei\" alese de autor sînt Solomon, Sulamita si corul fetelor din Israel. Interventiile lor se alterneaza mereu schimbînd planul actiunii si delimitînd succesiunea de 7 „tablouri\" idilice:\r\n\r\n1. Retrairea nuntii împaratesti (1:1-2:7)\r\n\r\n2. Amintiri din perioada începuturilor (2:8 - 3:5)\r\n\r\n3. Retrairea perioadei logodnei (3:6 - 5:1)\r\n\r\n4. Visul zbuciumat al miresei (5:2 - 6:3)\r\n\r\n5. Împaratul gîndindu-se la mireasa lui (6:4 - 7:10)\r\n\r\n6. Miresei îi este dor de acasa (7:11 - 8:4)\r\n\r\n7. O reînnoire a dragostei în Liban (8:5-14)\r\n\r\nSimbolistica „tipologica\" a acestei carti este înlesnita de existenta psalmului 45 cu care probabil ca a fost intentionata sa faca pereche. Va îndemnam sa cititi cu atentie acest psalm si sa vedeti cum, în spatele actiunii în care evolueaza Solomon si mireasa lui aleasa, se desfasoara un alt plan semantic în care Dumnezeu (Cristos) este mirele, iar Israelul (Biserica) este mireasa (conform talmacirii din Evrei 1:7, 8).\r\n\r\nContinutul psalmului 45 poate fi împartit în doua parti egale:\r\n\r\nI. O invocatie adresata mirelui (2-9)\r\na. despre frumusetea fapturii sale, v.2\r\nb. despre maretia înfaptuirilor lui, v. 3-5\r\nc. despre stabilitatea împaratiei lui, v. 6\r\nd. despre bucuria lui în fata casatoriei, v. 7-9\r\n\r\nII. O invocatie adresata miresei (10-17)\r\na. un apel la o daruire totala, v. 10-11\r\nb. promisiunea unei mari cinstirii, v. 12\r\nc. un elogiu al frumusetii ei, v. 13-15\r\nd. o promisiune a unui viitor stralucit, 16-17\r\n\r\nMesajul cartii: Pentru cititorul crestin, Cîntarea Cîntarilor constituie o subliniere a unirii noastre cu Cristos. În Biblie sînt si alte alegorii care desriu aceasta relatie. Se vorbeste despre Cristos ca si Cap, iar despre Biserica ca trup, definind aspectul de unire vie într-un organism al vietii; se vorbeste despre Cristos ca temelie si despre noi ca pietrele cladite în edificiul de deasupra, pentru a defini aspectul trainic al acestei uniri. Se spune ca Cristos este vita, iar noi sîntem mladitele, pentru a ilustra caracterul roditor al unirii noastre; în sfîrsit Cristos este prezentat ca „Cel dintîi nascut, dintre mai multi frati\" pentru a sublinia aspectul unei mosteniri comune pe care o avem prin unirea cu Cristos. Nici una dintre aceste alegorii nu este suficienta pentru a ilustra desavârsirea unirii noastre mistice cu Cristos, Mîntuitorul nostru. Numai unirea dintre un sot si o sotie, în contopirea totala din perimetrul familiei, poate reda ceva mai mult din aceasta extraordinara lucrare prin care Dumnezeu ne-a asezat „în Cristos\" (Efes. 5:31-32)." },
    //ISAIA
    { "ISAIA", "Explicatia Cartii: Isaia\r\n\r\nCeea ce este Shakespeare pentru literatura, Michelangelo pentru sculptura si Bach pentru muzica, aceea este Isaia pentru profetie. Calitatea scrisului sau se ridica deasupra tuturor celorlalti profeti ai Bibliei. Nu este de mirare ca profetia lui fara seaman a fost aleasa sa deschida seria celor 17 cartii profetice ale Vechiului Testament.\r\n\r\nConsiderat „Evanghelistul Vechiului Testament\", proorocul Isaia a fost instrumentul ales de Dumnezeu ca sa dea lumii cufundate în întunerec cea mai luminoasa carte de profetie Mesianica. „...Totusi întunericul nu va împarati vesnic pe pamîntul în care acuma este necaz...  Poporul care umbla în întunerec vede o mare lumina; peste cei ce locuiau în tara umbrei mortii rasare o lumina.\" (9:1, 2) „Caci un Copil ni s-a nascut, un Fiu ni s-a dat, si Domnia va fi pe umarul Lui; Îl vor numi: „Minunat, Sfetnic, Dumnezeu tare, Parintele vesniciilor, Domn al pacii\"(9:6) „De aceea Domnul însusi va va da un semn; Iata, fecioara va ramîne însarcinata, va naste un fiu, si-i va pune numele Emanuel (Dumnezeu este cu noi)\" (7:14)\r\nCartea poarta numele autorului ei, care tradus înseamna providential tocmai: „Mîntuirea este a Domnului\".\r\nStarea sociala a lui Isaia a fost înalta. Primit cu familiaritate de împaratii Ahaz si Ezechia (cap.7 si 37), el a fost cronicarul de la curte în timpul domniei lui Ozia si Ezechia (2 Cronici 26:22; 32:32) Cartea lui poarta pecetile unui om cu educatie aleasa: stil elegant, ritm, bogatie de pasaje cu o deosebita frumusete literara.\r\n\r\nNu stim nimic despre tatal sau, Amot, dar stim despre Isaia însusi ca a fost casatorit (7:3) si ca a avut doi copii: ªear Iasub („O ramasita se va întoarce\") si Maher-ªalal-Haz-Baz („Grabeste de pradeaza, arunca-te asupra prazii\") (8:3). Sotia lui a fost de asemenea proorocita. Este discutabil daca Isaia a fost din neam de preoti. Accesul lui în Templu s-ar parea ca indica acest lucru (6:6 comparat cu 2 Cronici 26:18)\r\n\r\nÎn privinta caracterului, putem spune ca Isaia a fost caracterizat de: (a) îndrazneala, deopotriva înaintea împaratilor si înaintea multimii, (b) patriotism înflacarat - este pornit împotriva a tot ceea ce ar putea aduce nenorocire poporului sau. Totusi, (c) blînd cu celelalte popoare din jur, într-o (d) simpatie care se ridica deasupra oricarui egoism national. Vorbeste cu (e) sarcasm si cu indignare împotriva pacatului, totusi este (f) atent în limbaj cînd vorbeste fiind plin de reverenta fata de Dumnezeu. Viata sa întreaga este plina de (g) spiritualitate. Pe Dumnezeu îl numeste adesea Cel prea înalt, Cel Sfînt, Cel Atotputernic. Toate aceste trasaturi îl fac pe Isaia unul dintre cele mai bune exemple pentru predicatorii din toate timpurile.\r\nTextul cartii plaseaza aceasta data: „Pe vremea lui Ozia, Iotam, Ahaz si Ezechia\" (1:1). Isaia a avut deci aproximativ 60 de ani de activitate profetica (740-680 î.Cr.). În vremea aceasta Israelul a fost dus în robie (722-721- î.Cr.), iar împaratia lui Iuda a fost partial invadata de ostile lui Sanherib (701 î.Cr.). Traditia evreiasca plaseaza moartea sa pe vremea lui Manase, care l-ar fi închis în trunchiul unui copac si l-ar fi taiat în doua cu ferastraul (Evrei 11:37)\r\nCei 60 de ani de activitate ai lui Isaia încep atunci cînd cele 10 semintii din regatul de Nord al lui Israel erau aproape sa fie cucerite de Asiria si duse în robie. Se încheiau astfel aproximativ 200 de ani de trista istorie în care Israelul ratacise departe de Iehova, sub conducerea a nu mai putin de 19 împarati din opt dinastii diferite.\r\n\r\nSub amenintarea lui Tiglat Pileser al Asiriei, Israelul intra în alianta cu Siria si Damascul. Împaratul Ahaz din Iuda refuza sa se alature acestei aliante si este tinta unei invazii de pedepsire (2 Regi 16 si 2 Cronici 28). Ahaz cere ajutorul puternicei Asirii, care zdrobeste Siria si Israelul, dar ia sub tutela sa si micul regat al lui Iuda. Vasalitatea împaratiei lui Iuda dureaza pîna cînd Ezechia se rascoala (2 Regi 18). La aceasta rascoala îl îndeamna Isaia care îl încurajeaza sa iasa din orice aliante omenesti si sa se încreada numai în Iehova. Speriat însa de amenintarea asiriana, Ezechia asculta de alte glasuri care-l îndemnau sa se alieze cu Egiptul, singura putere rivala pe masura Asiriei (Isaia 30:2-4). Cînd Sanherib, noul împarat al Asiriei, vine sa pedepseasca Iuda, Egiptul ezita sa trimeata ajutor, iar Ezechia este fortat sa se umileasca capitulînd si platind un mare pret în argint si aur (2 Regi 18:13-16). În secret însa, Iuda continua sa flirteze cu Egiptul. Aflînd aceasta, Sanherib se întoarce sa-si zdrobeasca toti rivalii. Din imensele trupe îndreptate înspre Egipt, desprinde o armata pe care o trimite sa cucereasca Ierusalimul (Isaia 36, 37). În panica asediului, Ezechia asculta de Isaia, striga catre Iehova si rezultatul nu întîrzie sa apara: Dumnezeu însusi se lupta cu asirienii producîndu-le o mare înfrîngere din care Sanherib nu si-a mai revenit niciodata, împaratia lui Iuda este eliberata astfel de orice amenintare si se va bucura pentru un timp de pace, liniste si prosperitate. Daca vom reusi sa tinem în minte acest mic rezumat istoric, citirea cartii lui Isaia va fi mai usoara, iar unele pasaje obscure vor putea fi cu usurinta clarificate.\r\nSubiectele rostirilor profetice ale lui Isaia acopar o arie impresionanta de teme. Ele se întind de la întîmplari petrecute în cer printre fapturile lui Dumnezeu înainte de creatia lumii (vezi Isaia 42:5) si pîna la evenimente din viitorul îndepartat în care Dumnezeu va crea ceruri noi si un pamînt nou (Isaia 65:17; 66:22). Chiar daca Isaia este important si din punct de vedere al profetiilor despre Ierusalim (pe care-l numeste în 30 de feluri diferite), despre Israel si Iuda sau despre celelalte natiuni ale lumii (Isaia 2:4; 5:26; 40:15, 17, 22; 66:18), totusi importanta lui majora ramîne aceea a unui crainic ce vesteste lucrarea lui Mesia, în textul profetiilor sale este scris clar despre evenimente care s-au întîmplat apoi întocmai în viata Domnului Isus Cristos. Isaia a scris despre nasterea Sa (7:17; 9:6), despre dumnezeirea Sa (9:6-7), despre lucrarea Sa (9:1-2; 42:1-7; 61:1-2), despre suferintele si moartea Sa\r\n\r\n(52:1-53:12) si despre împaratia Sa viitoare (cap. 2, 11, 65).\r\n\r\nCuvinte cheie si teme caracteristice: Dintre toate cartile Vechiului Testament, cartea lui Isaia este cea care vorbeste cel mai mult despre Domnul Isus Cristos. De fapt, personajul mesianic este numit „Robul Domnului\", între evrei si crestini exista o disputa asupra adevaratei lui identitati: evreii îl identifica cu Israelul, iar crestinii îl identifica cu Domnul Isus Cristos.\r\n\r\nEste evident ca profetia lui Isaia a avut si un sens local imediat destinat vremii si poporului spre care a fost rostit si în cadrul acesta „robul Domnului\" poate si trebuie sa fie identificat cu Israelul. Totusi, nimeni nu poate nega faptul ca profetia lui Isaia este si un mesaj peste veacuri, destinat celor cu care se vor împlini planurile viitoare ale lui Dumnezeu. În acest context al prevestirii, Robul Domnului trebuie identificat neaparat cu fiinta si lucrarea lui Isus Cristos din Nazaret, Mîntuitorul lumii. Combinînd cele spuse ajungem la concluzia ca în unele pasaje „Robul Domnului\" poate fi identificat cu Israelul, care a functionat ca un un „ante-tip\" al lui Cristos în lucrarea de mîntuire, în timp ce în alte pasaje este clar ca el nu poate fi identificat decît cu Mesia-Isus din Nazaret, Fiul lui Dumnezeu venit sa traiasca si sa moara pentru salvarea oamenilor.\r\n\r\nUn alt pasaj foarte cunoscut, dar nu mai putin controversat este Isaia 53. Biserica crestina a preluat acest capitol ca fiind al ei, fara sa mai tina cont ca el se adreseaza în primul rînd si în mod special neamului evreu. Nu Biserica, (care nu exista atunci) „nu L-a bagat în seama\" (53:3), ci Israelul! Ei fusesera acolo cînd a fost rastignit Robul Domnului si tot ei vor fi aceia care vor cînta în viitor aceasta „cîntare a plîngerii\", despre care ne vorbeste si profetia din Zaharia 12:10.\r\n\r\nÎntre Israel si Biserica continua si astazi acest conflict de interpretare. Evreii nu-L pot accepta astazi pe Domnul Isus-Cristos ca „Mîntuitor\" în planul lui Dumnezeu, dar nici unii din Biserica nu par a întelege ca Isaia ni-l prezinta solemn si pe Israel în rolul sau mesianic.\r\n\r\nPentru cititorul pasionat de desfasurarile politice din lumea contemporana, capitolele 60-66 sînt o lectura fascinanta. Textul vesteste reîntoarcerea evreilor în tara lor, refacerea statului Isarel, criza mondiala si framîntarile care nu se vor rezolva decît la revenirea lui Mesia.\r\nAranjarea cartii lui Isaia este usoara de tinut minte pentru ca ea oglindeste simetric Biblia. Ea are 66 de capitole exact cîte carti sînt si în Biblia noastra. Cartea lui Isaia, se împarte în doua parti cu exact acelasi numar de capitole cîte sînt si cartile din cele doua jumatati ale Bibliei: 39 de carti pentru Vechiul Testament si 27 de carti pentru Noul Testament.\r\n\r\nPrimele 39 de capitole ale cartii lui Isaia se ocupa ca si Vechiul Testament mai mult cu Legea, neascultarea si pedeapsa. Ultimele 27 de capitole vorbesc ca si Noul Testament mai mult despre har, mîntuire, restaurare si slava. Am putea spune ca profetia lui Isaia este o Biblie în miniatura, dovedind înca o data prin simetriile existente ca: „Toata Scriptura este insuflata de Dumnezeu\" (2 Timotei 3:16).\r\nI. PROFETIILE MÎNIEI 1 - 39\r\na. Judecata asupra lui Iuda, 1:1-31\r\nb. Ziua Domnului, 2:1-4:6\r\nc. Pilda cu via Domnului, 5:1-30\r\nd. Chemarea lui Isaia în slujba, 6:1-13\r\n\r\nASIRIA CUCEREªTE ISRAELUL\r\ne. Semnul lui Emanuel, 7:1-25\r\nf. Semnul lui Maher-ªalal-Has-Baz, 8:1-22\r\ng. Profetia despre venirea lui Mesia, 9:1-7\r\nh. Judecata lui Efraim, 9:8-10:4\r\n\r\nDUMNEZEU PEDEPSEªTE ASIRIA\r\nAsirianul smerit, 10:5-34\r\n\r\nÎMPaRatIA LUI MESIA 11-12\r\nJudecata împotriva neamurilor, 13:1-23:18\r\n\r\nVREMEA SFÎRªITULUI\r\na. Necazul cel mare, 24:1-23\r\nb. Lauda pentru împaratie, 25:1-12\r\nc. Cîntarea împaratiei, 26:1-21\r\nd. Propasirea lui Israel în împaratie, 27:1-13\r\ne. Cele sase „vai-uri\", 28:1-34:17\r\nf. Venirea împaratiei, 35:1-10\r\n\r\nSCURTa INCURSIUNE ISTORICa\r\na. Ezechia scapa de asirieni, 36:1-37:38\r\nb. Ezechia scapa de moarte, 38:1-22\r\nc. Pacatul lui Ezechia, 39:1-8\r\n\r\nII. PROFEtIILE MÎNGÎIERII 40 - 66\r\na. Mîngîiere prin izbavire, 40:1-11\r\nb. Mîngîiere prin Domnul, 40:12-41:29\r\nc. Mîngîiere prin Robul Domnului 42:1-44:27\r\nd. Mîngîiere prin bunavointa lui Cir, 44:28-45:25\r\ne. Mîngîiere prin pedepsirea Babilonului, 46:1-48:22\r\n\r\nPROFEtII DESPRE MESIA\r\na. Lucrarea lui Mesia, 49:1-26\r\nb. Ascultarea lui Mesia, 50:1-11\r\nc. Mesia încurajeaza Israelul, 51:1-52:12\r\nd. Mesia sufere pentru ispasire, 52:13-53:12\r\ne. Mesia promite refacerea lui Israel, 54:1-17\r\nf. Mesia cheama toate popoarele, 55:1-56:8\r\ng. Mesia avertizeaza pe cei rai, 56:9-57:21\r\n\r\nUN VIITOR GLORIOS PENTRU ISRAEL\r\na. Adevarata pocainta, 58:1-59:21\r\nb. Ierusalimul zidit din nou, 60:1-22\r\nc. Vestirea mîntuirii, 61:1-11\r\nd. Venirea mîntuirii, 62:1-12\r\ne. O zi de razbunare a Domnului, 63:1 -6\r\nf. Rugaciunea ramasitei, 63:7-64:12\r\ng. Dumnezeu raspunde ramasitei, 65:1-16\r\nh. Evenimentele glorioase de la sfîrsit, 65:17-66:24\r\nCauta in biblie\r\nTextul cautat\r\nCauta...\r\nCum caut\r\nToate Cuvintele\r\nUnde caut:\r\nToată biblia\r\nVersetul Zilei\r\nJudecatori 6:14-16\r\n\r\nDu-te cu puterea aceasta pe care o ai, si izbaveste pe Israel din mana lui Madian! oare nu te trimit Eu\r\nMeditatia Zilei\r\n\r\nPowered by Biblia Online\r\nFacebook\r\n\r\nParteneri\r\nStiri Crestine\r\nMuzica Crestina\r\nVersuri Crestine\r\nChristian Lyrics\r\nBiblia Online Cornilescu\r\nVersuri Crestine\r\nNu este Crăciun fără Isus - Estera & Laura Bretan\r\nÎn iesle azi - Diana Scridon Pop\r\nBucurie in suflet - Familia Timofte\r\nSe-aude Glas Peste Carpați / Tăria Noastră Fii Mereu - The Beuca Family\r\nUn singur Dumnezeu - Otto Pascal & Biji\r\nFiu iubit - BBSO\r\nClipa de clipa - Philadelphia Band\r\nPrintul Pacii a venit - Alin și Emima Timofte & TB Music Family\r\nPraise - Lumina Worship\r\nUnde ești țara mea?\r\nPowered by VersuriCrestine.ro\r\n" },
    //IEREMIA
    { "IEREMIA", "Explicatia Cartii: Ieremia\r\n\r\n„O! de mi-ar fi capul plin cu apa, de mi-ar fi ochii un izvor de lacrimi, as plînge zi si noapte pe mortii fiicei poporului meu!\" (Ieremia 9:1).\r\nCartea poarta numele autorului ei, care tradus înseamna: „Dumnezeu arunca\". Numele lui este un avertisment despre pedeapsa pe care profetul a fost trimis s-o vesteasca.\r\nDumnezeu l-a chemat în slujba pe Ieremia la aproximativ 60 de ani dupa moartea profetului Isaia. Activitatea lui avea sa se desfasoare pe durata ultimilor 40 de ani de existenta ai regatului lui Iuda.(„Cuvîntul Domnului i-a vorbit pe vremea lui Iosia... pe vremea lui Ioiachim... pîna la sfîrsitul lui Zedechia... pîna pe vremea cînd a fost dus Ierusalimul în robie\"(12:3).\r\n\r\nCînd Ieremia a început sa vorbeasca, la vîrsta de aproximativ 21 de ani, norii grei si amenintatori se ridicasera deja la orizontul istoriei lui Iuda. Samaria si întreg regatul de Nord cazusera deja prada de razboi si fusesera stramutati în tara amara a robiei. Idolatria copiilor lui Dumnezeu atrasese asupra lor judecata geloziei divine.\r\nSpre deosebire de cartea lui Isaia care ne spune foarte putin despre viata autorului ei, cartea proorocului Ieremia este presarata aproape peste tot cu pasaje de confesiune autobiografica (Ier. 10:23-24; 11:18 -12:6; 15:10-18; 17:9-11, 14-18; 18:18-23; 20:7-18)\r\n\r\n(a). Locul nasterii: Ieremia s-a nascut la Anatot, un satuc din tinutul lui Beniamin aflat la 4 km NE de Ierusalim. Numele locului s-a pastrat înca din vremurile vechi si deriva de la zeita feniciana Anat. Anatot era una din cele 13 cetati date Levitilor în teritoriile ocupate de Iuda, Simeon si Beniamin (Iosua 21:13-19; l Cronici 6:57-60). Dupa divizarea împaratiei lui Solomon, Anatot-ula ramas în împaratia lui Iuda. Localitatea mai exista si astazi sub numele de Anata.\r\n\r\n„Din tara lui Beniamin\". Asemenea marelui apostol Pavel de mai tîrziu si Ieremia a fost din semintia lui Beniamin si tot asemenea lui Pavel si el a primit o misiune îndreptata în acelasi timp si înspre iudei si înspre neamuri (Ieremia 1:5, 10, 18). Cu amîndoi s-au împlinit frumoasele cuvinte din promisiunea adresata de Dumnezeu cu secole înainte celor din semintia lui Beniamin:\r\n\r\n„El este preaiubitul Domnului, El va locui la adapost lînga Dînsul. Domnul îl va ocroti întotdeauna, si se va odihni între umerii Lui\" (Deuteronom 33:12).\r\n\r\nCu adevarat acesti doi beniamiti amenintati din toate partile au gasit odihna numai în Domnul, care a stiut sa-i ocroteasca si care i-a purtat cu credinciosie pe umeri.\r\n\r\n(b). Familia: Tatal sau s-a numit Hilchia (1:1) si a fost din rîndul preotilor. Ieremia a fost deci si preot si profet. Mama lui este mentionata în Ier. 15:10, dar nu i se aminteste numele. Ni se mai spune ca Ieremia a mai avut si alti frati (Ier. 12:6). Un amanunt semnificativ pentru consacrarea profetului este ca Dumnezeu nu i-a dat voie sa se casatoreasca si sa aiba copii (Ier. 16:2).\r\n\r\n(c). Chemarea lui Ieremia: „în al treisprezecelea an al domniei lui Iosia\" înseamna în anul 626 î.Cr. De fapt textul ne spune ca Dumnezeu l-a „plamadit\" pentru misiunea lui înca din pîntecele mamei (Ier. 1:5). Timiditatea si sfiiciunea lui proverbiala l-au facut însa sa primeasca numai cu greu misiunea care i-a fost încredintata (Ier. 1:5, 7, 8; 17:16; 20:7).\r\n\r\n(d). Moartea sa: Traditia ne spune ca Ieremia a murit în Egipt, în mijlocul ramasitei poporului, fiind omorît cu pietre în timp ce-i mustra pentru închinaciunea catre „împarateasa cerului\" (Ier. 44:1, 8, 16, 17, 18, 25, 26).\r\n\r\n(e). Caracterul profetului: Ieremia este unul dintre cele mai complexe si mai atragatoare caractere din galeria eroilor biblici. În launtrul lui, Dumnezeu a tesut într-o armonioasa întrepatrundere duiosia unei mame si statornicia unui luptator, tandretea feminina si hotarârea neînduplecata a unui barbat, fragilitatea nervoasa si simplicitatea transparenta, elocventa sensibila si duritatea proclamatorului de adevar.\r\n\r\nNatura sa launtrica este atît de vizibila în afara, convulsiile sufletului sau sînt atît de publice, încît îl putem asemana cu limpezimea apelor de clestar din lacurile montane care reflecta fidel turbulenta mereu schimbatoare a norilor de deasupra.\r\n\r\nIeremia a fost daruit de Dumnezeu cu o natura interioara care nu l-a lasat sa se restrînga la pozitia unui simplu „comunicator\" al vointei divine. El n-a fost niciodata capabil sa se detaseze afectiv de continutul mesajului care i-a fost încredintat. Mistuit de o dragoste intensa si chinuitoare, Ieremia si-a trait mesajele suferind si condamnînd deopotriva. Omul si discursul profetic s-au contopit în întregime.\r\n\r\nCe impresioneaza mai întîi cînd ne gîndim la Ieremia?\r\n\r\na. Simpatia cu care el sufere ca nevinovat alaturi de cei cazuti în vina. Este o simpatie si o identificare la intensitatea careia numai putini oameni au ajuns. Launtrul lui Ieremia era sfîsiat în doua. Pe de o parte, era îndragostit de Dumnezeu cu o iubire suprema, neclatinata si definitiva, iar pe de alta parte era îndragostit de concetatenii lui si nu se putea opri sa nu sufere alaturi de ei. Cînd erau loviti ei, el le simtea loviturile.\r\n\r\nPe de o parte, în relatia lui cu Dumnezeu, Ieremia era un profet, iar în relatia lui cu poporul era un patriot. Ieremia a reusit sa intre deopotriva si în viata poporului sau si în natura divina, identificîndu-se cu amîndoua. El nu s-a multumit sa vorbeasca „pentru\" Dumnezeu, ci a vorbit „împreuna\" cu El si nu s-a oprit doar sa vorbeasca poporului, ci s-a coborît sa fie „împreuna\" cu ei în suferinta.\r\n\r\nb. Perseverenta lui plina de rabdare. Dumnezeu i-a dat lui Ieremia una din cele mai imposibile misiuni. Remarcati situatia paradoxala în care a trait profetul: Dumnezeu l-a chemat sa vorbeasca unui popor razvratit, dar în acelasi timp i-a interzis sa mijloceasca pentru ei (Ier. 7:16; 14:11-12). L-a încredintat cu un mesaj care l-a facut urît de popor, dar în acelasi timp nu i-a dat nici o posibilitate de a iesi din mijlocul lor (Ier. 20:7-10 si mai ales 37:11-16 si 43:1-6). În toate acestea, cu lacrimi pe fata si cu focul în suflet, Ieremia a mers înainte, singuratic, predestinat sa nu aiba pe nimeni drag alaturi, lipsit de întelegerea unei sotii si de mîngîierea copiilor, urît de cei din jur, iubind fara masura, tratat ca tradator, dar punând în sîngele lui flacara nadejdii nationale, aruncat pe rînd în temnita, în gherla, în groapa cu noroi (Ier. 38:1-6), izbavit rînd pe rînd din toate acestea, nedorit ca prooroc si totusi tîrît cu forta alaturi de cei ce-si împlineau neascultarea fugind în Egipt.\r\nIeremia este „proorocul din ceasul al-12-lea\". El îl prezinta pe Dumnezeu ca fiind rabdator, dar plin de sfintenie. Cele doua lectii grafice din casa olarului ne spun ca un vas nereusit poate fi remodelat atîta timp cît nu s-a uscat înca (Ier. 18:1-4), dar daca se usuca, el nu mai poate fi recuperat si sfîrseste la groapa de gunoi (Ier. 19:10-11). Avertismentul dat de Dumnezeu a fost foarte clar: Timpul de asteptare pentru pocainta împaratiei lui Iuda se apropie curînd de sfîrsit. Din cauza împietririi inimii lor, robia Babiloniana a devenit inevitabila. Ieremia enumera cauzele morale si spirituale care atrag asupra natiunii catastrofa care urmeaza. El nu se opreste însa numai la aceasta. Dincolo de pedeapsa, el vesteste nadejdea într-o viitoare restaurare a Israelului. Privind în zarea vremii, el vede o ramasita care se va întoarce si va încheia cu Dumnezeu un nou legamînt (Ier. 31).\r\n\r\nCuvinte cheie si teme caracteristice: Cartea se desfasoara în jurul declaratiilor facute de Dumnezeu în Ier. 7:23, 24 si 8:11-12: „Caci iata porunca pe care le-am dat-o: „Ascultati glasul Meu...” Dar ei n-au ascultat... au dat înapoi si n-au mers înainte\". \"Leaga în chip usuratic rana poporului Meu, zicînd: „Pace! Pace!\" ªi totusi pace nu este. Vor fi dati de rusine, caci savîrsesc astfel de urîciuni... vor cadea împreuna cu cei ce cad, si vor fi rasturnati, cînd îi voi pedepsi, zice Domnul\".\r\nViata lui Ieremia este o încurajare pentru toti cei fragili si singuratici, pentru toti cei chemati de Dumnezeu, dar neprimiti de oameni. Cu rabdare, cu blîndete si cu perseverenta, el este purtat de Dumnezeu prin toate încercarile si reuseste sa-si duca pîna la bun sfîrsit slujirea. Multi din cei tari si înflacarati au cazut. Un chipes Saul, un întelept capabil ca Solomon, un viteaz ca Samson s-au rostogolit sub vînturile împotrivitoare ale vietii. Firavul si plîngaretul Ieremia a ramas însa în picioare. Uneori stejarii falnici cad sub apasarea furtunii, ramîn însa în picioare salciile plîngatoare unduite de vînt la marginea apelor.\r\nIntroducere - Chemarea lui Ieremia (1)\r\n\r\nI. PROFEtII GENERALE, NEDATATE 2-20\r\nPrimul mesaj 2:1- 3:5;\r\nal doilea mesaj 3:6-4:30;\r\nal treilea mesaj (la poarta Templului) 7:1-10:25;\r\ncel de al patrulea (legamîntul rupt) 11:1-12:17;\r\nal cincilea (semnul cu brîul de lîna) 13:1-27\r\nal saselea (despre seceta) 14:1-16:21;\r\nal saptelea (profetului necasatorit) 16:1-17:18;\r\nal optulea (la portie cetatii) 17:19-27;\r\nal noualea (vasul olarului) 18:1-23;\r\nal zecelea (vasul zdrobit) 19;\r\nurmarea 20.\r\n\r\nII. PROFEtII SPECIFICE, DATATE 19-39\r\nPrima (catre Zedechia) 21-23;\r\na doua (dupa prima deportare) 24;\r\na treia (despre robia Babiloneana) 25;\r\na patra (darîmarea Ierusalimului si a Templului) 26;\r\na cincea (la începutul domniei lui Ioiachim) 27-28;\r\na sasea (catre prinsii de razboi din Babilon) 29-31;\r\na saptea (al zecelea an al lui Zedechia) 32-33;\r\na opta (pe timpul asediului Babilonean) 34;\r\na noua (în zilele lui Ioiachim) 35;\r\na zecea (într-al patrulea an al lui hiachim) 36;\r\na unsprezecea (în timpul asediului) 37;\r\nurmarea 38-39.\r\n\r\nIII. DUPa CaDEREA LUI IUDA 40-44\r\nBlîndetea celor din Babilon fata de Ieremia, 40:1 -6;\r\nGhedalia ca dregator si omorîrea lui, 40:7-41:18;\r\nMesajul lui Ieremia pentru cei ramasi în tara 42;\r\nIeremia tîrît în Egipt 43:1-7;\r\nprimul mesaj catre iudeii din Egipt 43:8-13;\r\nal doilea mesaj catre cei din Egipt 44;\r\n\r\nIV. PROFEtII DESPRE NEAMURI 45-51\r\nUn mesaj introductiv adresat scribului credincios Baruc 54;\r\nprimul mesaj (despre Egipt) 46:1-28;\r\nal doilea (împotriva Filistenilor) 47:1-7;\r\nal treilea (împotriva Moabului) 48:1-47;\r\nal patrulea (împotriva Amonitilor) 49:1-6;\r\nal cincilea (împotriva lui Edom) 49:7-22;\r\nal saselea (împotriva Damascului) 49:23-27;\r\nal saptelea (împotriva Chedarului si Hatorului) 49:28-33;\r\nal optulea (împotriva Elamului) 49:34-39;\r\nal noualea (împotriva Babilonului si Caldeii) 50-51;\r\n\r\nÎNCHEIERE: - Cucerirea Ierusalimului 52." },
    //PLINGERILE LUI EREMIA
    { "PLÂNGERILE LUI IEREMIA", "Explicatia Cartii: Plangerile lui Ieremia\r\n\r\nCartea aceasta este o elegie scrisa la moartea unui oras. Ieremia plînge peste Ierusalimul transformat în ruine de catre ostile invadatoare ale Babilonului. Cartea poarta acelasi nume si în original si în traducerile mai cunoscute. Este clar ca durerea acestui profet l-a facut sa vibreze cu o maiestrie pe care numai un talent cu totul remarcabil putea s-o aiba. Cele cinci capitole ale cartii sînt tot atîtea plîngeri ale caror versete sînt asezate în ordine alfabetica. Fiecare dintre primele 4 capitole începe cu litera „a\" si se termina cu „z\". În capitolul 3, fiecarei litere îi sînt alocate un grup de trei versete. Avem de a face cu patru acrostihuri alfabetice. Ieremia plînge literalmente de la A la Z.\r\nNimeni nu putea sa scrie aceasta carte în afara lui Ieremia. Acest om a fost profetul unei iubiri neîmpartasite. Capitolele acestei carti ni-l prezinta plîngînd aplecat peste cadavrul marei sale iubiri: Ierusalimul. Aparent, misiunea lui Ieremia a fost un esec total. Dumnezeu nu ne judeca însa dupa rezultatele slujirii noastre, ci dupa credinciosia pe care am dovedit-o. Din acest punct de vedere, Ieremia este unul dintre cei mai mari barbati din istoria poporului lui Dumnezeu.\r\nCartea a fost scrisa la putin timp dupa distrugerea Ierusalimului (Ier. 39:52). Nebucadnetar a asediat orasul din luna Ianuarie 588 pîna în luna Iulie 586. Ierusalimul a cazut în 19 Iulie, 586, iar Templul si cetatea au fost date prada focului în 15 August 586. Probabil ca Ieremia a scris aceasta serie de elegii înainte de plecarea lui înspre Egipt (Ier. 43:1-7).\r\nAsa cum am spus, cele cinci capitole sînt tot atîtea plîngeri pe care profetul le revarsa asupra ruinelor orasului. Cartea aceasta este o extraordinara carte de poezie. Ea nu a fost asezata în sectiunea cartilor poetice pentru ca se înlantuie în lucrarea profetica a lui Ieremia, într-adevar, în mijlocul durerii si deznadejdii, Ieremia înalta din cînd în cînd ode ale sperantei, veritabile demonstratii de credinta în mijlocul celor mai adverse conditii. Iata, de exemplu, pasajul din capitolul 3:21-23: „Iata ce ma mai gîndesc în inima mea, si iata ce ma mai face sa trag nadejde: Bunatatile Domnului nu s-au sfîrsit, îndurarile Lui nu sînt la capat, ci se înnoiesc în fiecare dimineata. ªi credinciosia Ta este atît de mare!\".\r\n\r\nTrei teme se împletesc în desfasurarea tesaturii poetice:\r\n\r\n(1). Cea mai importanta dintre ele este bocetul pentru Ierusalimul darîmat. Dumnezeu s-a tinut de cuvînt. Lipsa pocaintei a atras asupra evreilor pedeapsa crunta. În durerea lui, Ieremia vorbeste cînd în numele sau, cînd în numele cetatenilor Ierusalimului, cînd în numele cetatii însasi.\r\n\r\n(2). A doua tema este marturisirea pacatelor si recunoasterea dreptatii lui Dumnezeu în hotarîrea pedepsei asupra împaratiei lui Iuda.\r\n\r\n(3) Cea de a treia tema este mai putin dominanta, dar nu mai putin importanta: ea este nadejdea într-o viitoare restaurare a Ierusalimului si a locuitorilor sai, datorita marelui har pe care evreii l-au gasit întotdeauna la Dumnezeul lor. Ieremia este un profet care cunoaste valoarea legamîntului. Iata ce scrie el în 3:24- 26: „Domnul este partea mea de mostenire\", zice sufletul meu; de aceea nadajduiesc în El. Domnul este bun cu cine nadajduieste în El, cu sufletul care-L cauta. Bine este sa astepti în tacere ajutorul Domnului\". ªi aceasta pentru ca -.„Domnul nu leapada pentru totdeauna. Ci, cînd mîhneste pe cineva, Se îndura iarasi de el, dupa îndurarea Lui cea mare\" (Plîngeri 3:31-32).\r\n\r\nCuvinte cheie si teme caracteristice: Explicatia nenorocirii Ierusalimului este descrisa astfel: „Ca un vrajmas a ajuns Domnul, a nimicit pe Israel, i-a darîmat toate palatele, i-a prapadit întariturile, si a umplut pe fiica lui Iuda cu jale si suspin. I-a pustiit cortul sfînt ca pe o gradina, a nimicit locul adunarii Sale; Domnul a facut sa se uite în Sion sarbatorile si Sabatul, si, în mînia Lui napraznica, a lepadat pe împarat si pe preot\" (Plîngeri 2:5-6).\r\nIeremia este un slujitor care se aseamana foarte mult cu Stapînul sau. Peste alte 600 de ani, Domnul Isus avea sa vina el însusi pe pamînt si sa plînga pentru aceiasi cetate: „Ierusalime, Ierusalime, care omori pe prooroci si ucizi cu pietre pe cei trimisi la tine! De cîte ori am vrut sa strîng pe copiii tai cum îsi strînge gaina puii sub aripi, si n-ati vrut! Iata ca vi se va lasa casa pustie\" (Matei 23:37-38).\r\nI. Distrugerea Ierusalimului (1)\r\nNenorocirile Ierusalimului 1:1-11\r\nJalea Ierusalimului 1:12-22\r\n\r\nII. Mînia lui Dumnezeu (2)\r\nJudecata Domnului, 2:1-10\r\nPlîngerea profetului, 2:11-22\r\n\r\nIII. Rugaciune pentru îndurare (3)\r\nCîntarea de jale 3:1-18\r\nPricina de încredere, 3:19-42\r\nSuferinta profetului 3:43-54\r\nRugaciunea profetului, 3:55-66\r\n\r\nIV. Jalea Ierusalimului asediat (4)\r\nAsedierea cetatii, 4:1-12\r\nMotivatia nenorocirii, 4:13-20\r\nO geana de nadejde, 4:21-22\r\n\r\nV. Rugaciune pentru renastere (5)\r\nMarturisirea, 5:1-18\r\nMijlocirea, 5:19-22" },
    //EZECHEL
    { "EZECHEL", "Explicatia Cartii: Ezechiel\r\n\r\n„Fiul omului, te pun pazitor peste casa lui Israel! Cînd vei auzi un cuvînt care iese din gura Mea, sa-i înstiintezi din partea Mea\" (Ezec. 3:16-21).\r\n\r\nIsaia este profetul Fiului lui Dumnezeu; Ieremia este purtatorul de cuvînt al Tatalui, care-si disciplineaza copiii; Ezechiel este proorocul Duhului, numit în text: „slava Domnului\".\r\nCartea poarta numele autorului ei care se poate traduce prin „Dumnezeu întareste\" sau „Domnul este taria mea\".\r\nCa si Isaia, Ezechiel a fost si preot si profet. El se pregatea sa slujeasca Domnului la Templul din Ierusalim, dar nu a mai avut ocazia s-o faca. Preotii erau instalati în slujire numai de la vîrsta de 30 de ani, iar Ezechiel a fost dus în robia babiloniana în anul 597 î.Cr., pe cînd avea numai 25 de ani. Dupa înca cinci ani petrecuti în Babilon, Ezechiel si-a început slujirea printre poporul captiv de la rîul Chebar. Timp de cel putin 22 de ani activeaza ca preot si profet (Ezec. 29:17). Se casatoreste, dar nu stim daca a avut sau nu copii din aceasta casatorie.\r\nAmanuntele date de Ezechiel ne ajuta sa stabilim foarte usor data scrierii. Fiecare sectiune profetica a cartii debuteaza cu data la care a fost primita de la Domnul si transmisa oamenilor. Activitatea profetica a lui Ezechiel a început în luna Iulie din anul 593 î.Cr. si a continuat cel putin pîna la vremea mentionata în ultima profetie înregistrata de textul cartii, adica: luna Aprilie din anul 571 î.Cr. Ezechiel a fost contemporan cu Ieremia si Daniel. Cam de aceiasi vîrsta cu Daniel, dar mult mai tînar decît Ieremia, profetul a fost mult influentat de „profetul lacrimilor” caruia i-a continuat aparent trei dintre profetiile sale: (1) vedenia cazanului pus pe foc (Ezec. 11:1-12; 24:3-14; comparate cu Ier. 1:13-15), (2) zicala strugurilor acrii (Ezec. 18:2-32; comparat cu Ier. 31:29-30) si (3) pilda celor doua surori (Ezec. 23:1-49; comparat cu Ier. 3:6-11).\r\nEzechiel ajunge la Babilon cu 9 ani dupa Daniel, facînd parte din cei 10.000 de captivi luati de Nebucadnetar în timpul domniei lui Zedechia (2 Împ. 24:11-18). El îsi începe activitatea dupa 5 ani de robie, spulberînd chiar de la început nadejdiile false ale celor care credeau ca robia va fi scurta si ca evreii se vor întoarce foarte repede acasa. Cartea lui Ezechiel a fost un mesaj adresat natiunii. Evreii trebuiau convinsi ca nu poate fi vorba despre o întoarcere în tara promisa, mai înainte de a se produce o totala întoarcere la Dumnezeu. Acest mesaj îsi atinge punctul maxim în pasajul din Ezec. 18:30-32:\r\n\r\n„Întoarceti-va si abateti-va de la toate faradelegile voastre, pentru ca sa nu va duca nelegiuirea la pieire. Lepadati de la voi toate faradelegile, prin care ati pacatuit, faceti-va rost de o inima noua si un duh nou. Pentru ce vreti sa muriti, casa a lui Israel? Caci Eu nu doresc moartea celui ce moare, zice Domnul Dumnezeu. Întoarceti-va dar la Dumnezeu, si veti trai\".\r\n\r\nCartea lui Ezechiel are un caracter autobiografic si este aranjata într-o ordine cronologica. Continutul cartii înregistreaza activitatea lui Ezechiel ca profet. Proorociile rostite de Ezechiel pot fi rezumate si împartite în felul urmator: (a) profetii adresate situatiei prezente: judecata Ierusalimului si a poporului evreu (Ezec. 4-14), (b) profetii despre viitor - destinul natiunilor (Ezec. 25-39) si (c) profetii despre vremea sfîrsitului: Templul, închinarea, cetatea (Ezec. 40-48).\r\n\r\nEzechiel este un prooroc foarte putin citat în Noul Testament. Nu-l gasim în textul Evangheliilor sau al epistolelor, dar el apare în toata maretia lui în cartea Apocalipsei. Pentru cel ce cunoaste Biblia, nu este nici un secret ca Ezechiel si Ioan sînt amîndoi profeti apocaliptici. ªi unul si celalalt au avut harul sa primeasca de la Domnul descoperiri extraordinare despre vremurile sfîrsitului, în care Dumnezeu va face ceruri noi si un pamînt nou în care va locui neprihanirea. Multe din pasajele cartilor lor pot fi citite si studiate în paralel. Iata de exemplu viziunea scaunului de domnie al lui Dumnezeu:\r\n\r\n„Deasupra cerului care este peste capetele lor, era ceva ca o piatra de safir, în chipul unui scaun de domnie; pe acest chip de scaun de domnie se vedea ca un chip de om, care sedea pe el\" (Ezec. 1:26).\r\n\r\n„Numaidecît am fost rapit în Duhul. ªi iata ca în cer era pus un scaun de domnie, si pe scaunul acesta de domnie sedea Cineva. Cel ce sedea pe el, avea înfatisarea unei pietre de iaspis si de sardiu; si scaunul de domnie era înconjurat cu un curcubeu ca o piatra de smaragd la vedere\" (Apoc. 4:2-3).\r\n\r\nCa si cartea lui Ioan, si proorociile lui Ezechiel sînt pline de vedenii (Ezec. 8), pilde (Ezec. 17), poeme (Ezec. l9), proverbe (Ezec. 12:22, 23; 18:2) si tablouri cu încarcatura simbolica. Specificul acesta mai are înca o explicatie. Pîna la caderea Ierusalimului, Dumnezeu l-a facut literalmente pe Ezechiel mut (Ezec. 3:26, 27). Asa ca mesajul lui catre popor nu a putut fi transmis prin vorbe, ci prin intermediul gesturilor si al simbolurilor. Din cauza acestei situatii, profetul a fost împins de cîteva ori de Dumnezeu în situatii dificile. A trebuit sa umble gol, sa stea legat cu frînghii, sa manînce baliga de animal, etc. Cel mai greu de purtat dintre toate semnele profetice a fost însa moartea sotiei sale, produsa chiar în ziua în care a cazut Ierusalimul (Ezec. 24:15-27). Pentru ca glasul lui Ezechiel catre popor sa sune la fel cu glasul Dumnezeului lor, Domnul a vrut ca profetul sa guste ceva din durerea pe care Dumnezeu însusi a simtit-o cînd a încuviintat pedepsirea cetatii favorite.\r\n\r\nUnul dintre cele mai grafice capitole ale cartii lui Ezechiel este capitolul 37 în care gasim viziunea vaii oaselor. Nicaieri în Biblie nu este descrisa mai plastic refacerea si renasterea poporului evreu. ªi tot nicaieri în Biblie nu se gaseste mai clar declaratia ca Dumnezeu nu si-a terminat planurile cu evreii. Dumnezeu poate sa-i ridice chiar si din morminte pentru a-i aseza în matca destinului lor mesianic.\r\n\r\nCuvinte cheie si teme caracteristice: Expresia caracteristica a cartii este „slava Domnului\". Ea apare de 12 ori în primele 11 capitole, dispare apoi din text, pentru ca sa reapara în capitolul 43 al cartii. Capitolul l debuteaza cu descrierea slavei ceresti, iar capitolele 40-48 încheie cartea cu descrirea instaurarii slavei lui Dumnezeu pe pamînt („Tatal nostru care esti în ceruri! Sfinteasca-se Numele Tau; faca-se voia Ta; vie împaratia Ta; precum în cer asa si pe pamînt\" - Matei 6:9-10). Între aceste doua extreme, cartea lui Ezechiel ne prezinta ce s-a întîmplat cu slava Domnului în împaratia lui Iuda. În Ezechiel 8 îl vedem pe profet dus de Domnul la Ierusalim ca sa fie martor la idolatria poporului si pentru a asista la plecarea slavei Domnului din cetate.\r\n\r\nÎn Vechiul Testament, slava Domnului era semnul prezentei lui Dumnezeu în mijlocul poporului. Locul ei era între heruvimii harului de pe chivotul legamîntului din Cortul întîlnirii. Din pricina pacatelor poporului, în timpul vietii lui Ezechiel, Dumnezeu trimite asupra evreilor pedeapsa robiei. Înainte însa de caderea Ierusalimului, slava Domnului paraseste treptat cetatea. Mai întîi, ea „s-a ridicat dintre heruvimul pe care era si s-a îndreptat spre pragul casei\" (Ezec. 9:3). Aceasta miscare a ei a facut ca Templul sa se umple de nor si curtea de stralucire (Ezec. 10:4). Apoi „slava Domnului a plecat din pragul Templului, si s-a asezat pe heruvimi. Heruvimii si-au întins aripile... si s-au oprit la intrarea portii Casei Domnului spre rasarit\" (Ezec. 10:18-19). „Dupa aceea... slava Domnului s-a înaltat din mijlocul cetatii, si s-a asezat pe muntele de la rasaritul cetatii\" (Ezec. 11:22-23). Plecarea slavei Domnului din Ierusalim s-a facut cu o încetineala si cu o maiestate care subliniaza parca regretul cu care Dumnezeu trebuie sa se desparta de cetate si sa îngaduie distrugerea ei. Cînd slava a plecat, a început prapadul.\r\n\r\nEzechiel nu se opreste însa aici. Profetia lui ridica valul de la orizontul viitorului. Capitolele finale ale cartii lui proclama întronarea împaratiei lui Israel si revenirea slavei: „M-a dus la poarta, la poarta dinspre rasarit. ªi iata ca slava Dumnezeului lui Israel venea de la rasarit... Slava Domnului a intrat în Casa pe poarta dinspre rasarit. Atunci, Duhul m-a rapit si m-a dus în curtea dinlauntru. ªi Casa era plina de slava Domnului!\" (Ezec. 43: l -5).\r\nNu trebuie sa cautam prea mult pentru a descoperi care este ideea principala si mesajul central din cartea lui Ezechiel. Ele se gasesc pe aproape fiecare fila a cartii. Cu variatii minime, fraza: „ªi vor sti ca Eu sînt Domnul\" se repeta de nu mai putin de 70 de ori. Ea este folosita de 29 de ori atunci cînd se vorbeste despre pedeapsa pe care o va aduce Iehova asupra Ierusalimului; de 24 de ori atunci cînd se vorbeste despre judecata divina care va cadea asupra neamurilor; si de 17 ori atunci cînd se vorbeste despre viitoarea restaurare a Israelului si despre binecuvîntarea care va veni asupra lui la vremea sfîrsitului. Aceasta observatie da pe fata însasi chintesenta cartii. Ezechiel anunta ca poporul ales, si toate celelalte popoare ale lumii, vor ajunge sa cunoasca în scurgerea vremii ca Iehova este singurul si adevaratul Dumnezeu, Suveran absolut asupra neamurilor si asupra istoriei.\r\n\r\nAceasta se va face prin trei actiuni distincte:\r\n\r\n(a), mai întîi este vorba despre pedepsirea Ierusalimului si despre ducerea poporului ales în robie, lucru despre care noi stim astazi ca s-a întîmplat întocmai;\r\n\r\n(b), în al doilea rînd, prin judecata si pedepsirea popoarelor existente în zilele lui Ezechiel, lucru despre care iarasi noi stim ca s-a întîmplat întocmai; si\r\n\r\n(c), în al treilea rînd, prin prezervarea Israelului si prin finala lui restaurare în pozitia si privilegiile de popor al Legamîntului. Aceasta s-a întîmplat partial o data cu reîntoarcerea din robia Babiloneana a „ramasitei\" sub conducerea lui Ezra si Neemia, se întîmpla acum prin supravietuirea miracuIoasa a Israelului si îsi va avea curînd încununarea în instaurarea împaratiei Mielului.\r\n\r\nRedus la o singura afirmatie, cartea lui Ezechiel spune: „ªi veti cunoaste ca Eu sînt Domnul!\".\r\nViziunea initiala si chemarea profetului Ezechiel în slujba (1-3)\r\n\r\n1. PREZENT - JUDECATA IERUSALIMULUI ªI A POPORULUI (4-24)\r\nPilde si profetii despre apropierea pedepsei (4-7)\r\nViziunea Templului si a cetatii: Slava Domnului pleaca (8-11)\r\nAlte simboluri si mesaje despre judecata (12-24)\r\n\r\n2. VIITOR - DESTINUL NAtIUNILOR (25-39)\r\nJudecati asupra popoarelor straine (25-32)\r\nDupa judecata Israelul va fi restaurat (33-37)\r\nDistrugerea lui GOG si MAGOG; înaltarea Israelului (38-39)\r\n\r\n3. VREMEA SFÎRªITULUI: TEMPLUL, ÎNCHINAREA, CETATEA (40-48)\r\n\r\nTEMPLUL recladit si slava cea noua (40:1 - 43:12)\r\nÎNCHINAREA reînnoita, rîul cel sfînt (43:13- 47:12)\r\ntARA re-împartita si\r\nCETATEA lui Dumnezeu (47:13- 48:35)" },
    //DANIEL
    { "DANIEL", "Explicatia Cartii: Daniel\r\n\r\nÎntreaga Biblie poate fi asemanata cu o oglinda plimbata asupra istoriei omenirii. Reciproca este valabila, caci întreaga istorie umana poate fi desprinsa din continutul profetic asezat în cele 66 de carti ale Bibliei.\r\n\r\nCartea lui Daniel poate fi numita pe drept cheia descifrarii mesajelor profetice. Cadranul unui ceas n-ar fi de nici un folos fara ajutorul limbilor indicatoare. Tot asa, cartea lui Daniel este pentru întregul Scripturii, indicatorul care ne ajuta sa ne orientam în succesiunea de evenimente consemnate în profetii.\r\n\r\nDin cauza continutului ei extraordinar, cartea lui Daniel a fost contestata de toti cei care au combatut Biblia si credinta crestina. Criticii ei înfierbîntati au prezentat-o drept un fals istoric produs cel mai devreme prin jurul anului 164 î.Cr. cu scopul de a-i întari în credinta pe evreii care treceau prin vremuri grele sub conducerea Macabeilor.\r\n\r\nDescoperirea manuscriselor de la Marea Moarta a redus la tacere aceste critici. Pentru crestini, lucrurile erau clare chiar si înainte de aceste descoperiri, în afara dovezilor continute chiar în textul cartii, despre Daniel a mai scris si Ezechiel, care-l aseaza în rîndul sfintilor poporului evreu, numindu-l: „un om neprihanit\" si „un om foarte întelept\" (Ezecniel l:19-20; 28:3) Nimeni nu s-a îndoit vreodata ca Ezechiel a trait în vremea cînd ni se spune ca a trait, ori el îl aminteste pe Daniel ca pe un contemporan al sau aflat înca în viata.\r\n\r\nCea mai mare dovada despre veridicitatea continutului cartii lui Daniel vine însa din partea Domnului Isus care citeaza din cartea profetului de trei ori în cuprinsul cuvântarii Sale din Matei 24. În versetul 15, El citeaza Daniel 8:13; 9:27; 11:31; 12:11; indicîndu-le ucenicilor semnele care-i vor anunta ca e timpul sa fuga din Ierusalim („De aceea, cînd veti vedea \"urîciunea pustiirii\", despre care a vorbit proorocul Daniel, „asezata în locul sfînt\" - cine citeste sa înteleaga! -...”)\r\n\r\nÎn versetul 21, El descrie venirea „Necazului cel mare\" citînd din Daniel 12:1: „Pentru ca atunci va fi un necaz asa de mare, cum n- a fost niciodata de la întemeierea lumii pîna acum si nici nu va mai fi\".\r\n\r\nApoi, în versetul 30, El descrie cea de a doua venire a Sa folosind textul din Daniel 7:13: „Atunci se va arata în cer semnul Fiului omului...”\r\n\r\nMai mult decît atît, în cel mai solemn moment din cadrul procesului religios la care a fost supus, în ceasul în care Marele Preot l-a întrebat categoric: „Te jur pe Dumnezeul cel viu sa ne spui daca esti Cristosul, Fiul lui Dumnezeu\". Domnul Isus i-a raspuns cu acelasi text din Daniel 7:13, 14: „Da, sînt. Ba mai mult, va spun de acum încolo veti vedea pe Fiul omului sezînd la dreapta puterii lui Dumnezeu, si venind pe norii cerului\" (Matei 26:63-64).\r\n\r\nDomnul Isus îl prezinta pe Daniel drept „un crainic al lucrurilor viitoare\" confirmînd ceea ce scrisese însusi Daniel despre cartea sa: „Tu însa Daniele, tine ascunse aceste cuvinte si pecetluieste cartea, pîna la vremea sfîrsitului. Atunci multi o vor citi si cunostinta va creste\" (Dan. 12:4). La venirea Sa Domnul Isus a proclamat începutul acestei „vremi a sfîrsitului\", a ridicat pecetea de pe cartea lui Daniel si a declarat emfatic: (De acum) „Cine citeste sa înteleaga\" (Matei 24:15).\r\n\r\nAcestei întelegeri îi dedicam rîndurile care urmeaza.\r\nCartea poarta numele autorului ei: Daniel, care se poate traduce prin „Dumnezeu este judecatorul\".\r\nDesi pare surprinzator, stim foarte putine lucruri despre acest om extraordinar.\r\n\r\n(1). Ca tînar, a fost dus ca rob din Israel la Babilon, în robie a hotarît sa nu-si piarda specificul national evreiesc, refuzînd sa „se spurce mîncînd din bucatele aduse de la masa împaratului\" (Dan. 1:8). Pus în scoala, a dat curînd dovada de o întelepciune cu totul iesita din comun, dar mai ales s-a evidentiat prin capacitatea supranaturala de a patrunde în lumea lui Dumnezeu si de a primi talmaciri pentru vise si vedenii.\r\n\r\n(2). În urma simpatiei de care s-a bucurat din partea lui Nebucadnetar, a fost promovat în cele mai înalte dregatorii imperiale. ªi acolo însa si-a pastrat dorinta de a trai ca un evreu cucernic. Din aceasta cauza, la vîrsta de 70-75 de ani a fost aruncat într-o groapa cu lei, din care Dumnezeu l-a scapat în chip miraculos. Caracterul sau integru si apartenenta lui la lumea supraomeneasca a revelatiei divine, l-au ajutat sa pluteasca asemenea unei corabii peste talazurile framîntate ale istoriei.\r\n\r\nMonarhii lumi s-au succedat unul dupa altul: Nebucadnetar a fost urmat de Belsatar, iar apoi Babilonul a cazut în stapînirea lui Dariu, medul si a lui Cir, persanul. Simpatizat de toti acestia, Daniel a ramas mereu la curtea împarateasca, harazit de Dumnezeu sa fie un fel de „crainic dumnezeiesc\" pe lînga curtile imperiale ale pamîntului. La ceas de cumpana, el a fost, asemenea poporului ales din care a facut parte, un fel de constiinta cosmica si de „lumina a neamurilor\".\r\n\r\nA murit la Susa, probabil de batrînete, la 90-94 de ani.\r\nUnul dintre motivele pentru care cartea lui Daniel a fost contestata este si continutul supranatural al celor scrise de el. Vise, vedenii, talmaciri, izbaviri miraculoase din foc sau din gura leilor, aparitii angelice sînt întrepatrunse într-o tesatura densa si imposibil de separat.\r\n\r\nDe fapt exista un motiv foarte întemeiat pentru care factorul supranatural este atît de proeminent în aceasta carte, în vremea aceea Israelul se afla în captivitate. Ierusalimul se gasea în ruina. Chiar si Templul - ultima speranta a evreilor - fusese ras de pe fata pamîntului. Într-un fel, Iehova, Dumnezeul evreilor se dovedise mai slab decît dumnezeii Babilonului! Bel-Merodah îl învinsese pe Iehova, sau cel putin asa le placea sa creada celor din pagînul imperiu. ªi tot asa erau înclinati sa creada si evreii. Reîntoarcerea si refacerea pareau imposibile. Nimeni nu credea în spusele lui Ieremia care estimase durata robiei la numai 70 de ani. La urma urmei, de ce s-ar tine Dumnezeu de aceasta promisiune daca nu s-a tinut de promisiunile facute lui David si Solomon?!\r\n\r\nLucrarile supranaturale relatate de Daniel au fost tocmai niste raspunsuri date unor astfel de gînduri. Ele au avut caracterul unor semne venite din partea lui Iehova si adresate lui Israel si neamurilor deopotriva.\r\n\r\nAtunci cînd Dumnezeu a transferat dreptul de instrument al suveranitatii Sale de la Israel si l-a daruit lui Nebucadnetar, El l-a ridicat pe evreul Daniel si l-a asezat la curtea regala Babiloneana pentru ca prin buzele lui si prin actiunile lui sa-l învete pe Nebucadnetar despre necesitatea supunerii înaintea Celui Atotputernic, încetînd sa mai vorbeasca la Ierusalim, Dumnezeu si-a plasat solul la curtea marilor imperii ale lumii, dovedind neamurilor ca El exista si ca are ultimul cuvînt în desfasurarea evenimentelor istoriei lumii. În acelasi timp, prin activitatea lui Daniel, Dumnezeu a dovedit evreilor ca El continua sa le fie un Dumnezeu „de aproape\", gata sa le asigure supravietuirea si în stare sa-i elibereze si acum din robie tot asa cum a facut-o si atunci cînd i-a scos din tara Egiptului.\r\nCartea lui Daniel este plina de întîmplari în care supranaturalul invadeaza lumea obisnuita si schimba cursul istoriei. A fost vremea cînd poporul lui Dumnezeu a avut nevoie de o revelatie deosebita, cu puternice implicatii în desfasurarea ulterioara a istoriei.\r\n\r\nCea mai buna dovada ca aceste evenimente supranaturale s-au petrecut întocmai si nu au fost produsul fabulatiei o constituie chiar transformarea uluitoare pe care ele le-au produs asupra poporului evreu.\r\n\r\nPe durata robiei, evreii si-au schimbat total atitudinea religioasa, devenind dintr-un popor înclinat spre idolatrie, un popor cu o credinta monoteista mai tare decît granitul. Cei plecati în robie s-au aflat acolo din cauza celor aproximativ 500 de ani de cochetarie cu idolii neamurilor învecinate. În numai 70 de ani petrecuti în captivitate, Dumnezeu i-a „curatit\" pe evrei în cuptorul mîniei Lui si i-a scos de acolo vindecati pentru vecie. Nu multimea idolilor din Babilon i-a schimbat pe evrei, ci tocmai acele manifestari dumnezeiesti supranaturale amintite în cartea profetului Daniel.\r\n\r\nDe la Ezechiel stim ca Daniel devenise celebru chiar fiind înca în viata. ªi cum s-ar fi putut sa fie altfel dupa întîmplari ca cele relatate în capitolele 2 si 3 ale cartii sale? Ba înca sub influenta lui Daniel, Nebucadnetar emisese si celebrele lui proclamatii catre toate popoarele din Imperiu în care recunostea suveranitatea Dumnezeului evreilor (cap.4). Asemenea lucruri nu puteau sa-i lase indiferenti pe cei din Israel. Cu o noua si arzatoare dorinta, ei s-au aplecat asupra scrierilor profetice ale lui Ieremia („în anul dintîi al domniei lui, eu, Daniel, om vazut din carti ca trebuiau sa treaca 70 de ani pentru darîmaturile Ierusalimului, dupa numarul anilor despre care vorbise Domnul catre proorocul Ieremia\" - cap.9:2), si ale lui Isaia. Acesta vorbise nu numai despre caderea viitoare a Babilonului, dar pomenise pe nume chiar si pe cel ce avea sa dea decretul de reîntoarcere a iudeilor în Canaan si de rezidire a templului din Ierusalim: Cir persanul (Isaia 45 si 46).\r\n\r\nCum trebuie sa fi trait evreii toate aceste evenimente? Cum trebuie sa se fi uimit ei vazîndu-le prinzînd viata chiar sub ochii lor? Experientele din acesti 70 de ani de evenimente supranaturale explica spulberarea îndoielilor si ratacirilor lor si transformarea lor într-un popor cu o adorare totala, unica si definitiva pentru Iehova.\r\n\r\nAranjarea cartii este foarte clara: primele 6 capitole sînt istorice, iar ultimele 6 capitole sînt profetice.\r\nScopul cartii exprimat în mesajul ei central este expus în cuvintele repetate emfatic de trei ori în cuprinsul capitolului 4 (4:17, 25, 32): „Ca sa stie cei vii ca Cel Prea înalt stapîneste peste împaratia oamenilor, ca o da cui îi place\". Este semnificativ ca acest mesaj este facut sa ajunga la noi prin gura smeritului Nebucadnetar, „capul\" de aur si cel dintîi suveran mondial din acesta „vreme a neamurilor\".\r\n\r\nO alta caracteristica izbitoare a acestei carti este aceea ca ea este scrisa în doua limbi. De la capitolul 2:4 pîna la capitolul 7 este scrisa în limba aramaica. Restul textului este scris însa în limba ebraica. Are faptul acesta vreo semnificatie? Noi credem ca da.\r\n\r\nExista o corespondenta extraordinara între visul dat de Dumnezeu lui Nebucadnetar (cap.2) si prima vedenie a lui Daniel (cap.7). Amîndoua ne traseaza cursul general al istoriei din „vremea neamurilor\". Celelalte vedenii ale lui Daniel privesc înspre viitor mai ales din punct de vedere al poporului Israel. Spre a le separa de restul, capitolele de la 2 la 7 sînt scrise în aramaica, limba care se vorbea atunci în institutiile comerciale si diplomatice ale lumii. Aceasta schimbare, de la limba ebraica din debutul cartii, la limba aramaica din mijlocul ei si apoi întoarcerea la limba ebraica ilustreaza schimbarea accentului pe care Dumnezeu îl va pune în istoria lumii si mai spune ceva, si anume, ca Dumnezeu a dat lumii acces numai la planul general al vremurilor, pastrînd doar pentru cei ce stiau limba evreilor detaliile lui semnificative.\r\n\r\nÎn plus, folosirea ambelor limbi este înca o dovada ca Daniel si-a scris cartea tocmai în acea vreme. Înainte de vremea robiei, evreii n-ar fi înteles aramaica (vezi 2 Regi 18:26), iar dupa robia babiloneana, ei n-ar mai fi înteles ebraica, deoarece încetasera sa o mai foloseasca (vezi Neemia 8:8). Numai si numai în perioada de timp în care a trait Daniel, evreii au cunoscut amîndoua limbile amintite.\r\n\r\nDumnezeu a asezat chiar în cuprinsul cartii dovada datei evenimentelor anuntate de autorul acestei scrieri. Astfel, El i-a redus la tacere pe cei care vor sa-i conteste valabilitatea, pentru ca de fapt nu vor sa-i primeasca mesajul.\r\n\r\nCuvinte cheie si teme caracteristice: Pentru studiul nostru, cel mai important este felul în care Domnul Isus îl remarca pe profetul Daniel drept un crainic al lucrurilor viitoare (Matei 24:15-16). Datele prezentate în cartea lui Daniel sînt într-atît de exacte si de clare încît mai marii iudeilor au interzis cu desavîrsire studierea lor, ca nu cumva poporul sa priceapa ca Isus a fost într-adevar „Unsul-Mesia\".\r\n\r\nMulti alti critici au sustinut cu încapatînare ca amanuntele istorice au fost adaugate mult mai tîrziu în cartea lui Daniel de catre preoti. Descoperirile de la Marea Moarta (între altele si o copie a cartii lui Daniel) au dovedit însa zdrobitor de clar ca textul profetic din care a citat Domnul Isus a existat exact în forma de azi cu sute de ani înainte de era crestina.\r\n\r\nCartea lui Daniel este cheia profetica a Bibliei. Imaginea ei se reflecta simetric în cartea Apocalipsei. Cele mai importante pasaje profetice ale cartii sînt visele profetice ale lui Nebucadnetar din capitolul 2 si vedenia despre „cele saptezeci de saptamîni\" din capitolul 9.\r\n\r\nVisul lui Nebucadnetar din capitolul 2 schiteaza istoria împaratiilor lumii si are un interes deosebit pentru neamuri. Vedenia lui Daniel din capitolul 9 este o profetie privitoare la Israel si la evenimentele cheie prin care va trece acesta.\r\n\r\nVisul lui Nebucadnetar\r\n\r\nNiciodata nu a visat un muritor un vis mai epocal decît acesta. Împreuna cu vedenia lui Daniel din cap.7 (care reia tema cap.2 si o lamureste în detaliu), acest pasaj anunta ca Dumnezeu s-a hotarît sa renunte pentru un timp la importanta preponderenta pe care a jucat-o în istorie divina Israelul si sa aduca în scena \"Vremea neamurilor\", pomenite de Domnul Isus în Luca 21:24.\r\n\r\nLui Nebucadnetar, preocupat de „ce se va întîmpla în vremurile de pe urma\" (Dan. 2:28), Dumnezeu îi descopera în vis succesiunea de puteri mondiale care se va desfasura pe pamînt începînd cu Imperiul lui si pîna la venirea împaratiei „care va dainui vesnic\" (Dan. 2:44.) Aceasta ultima împaratie pe care Dumnezeu o va aduce pe pamînt va fi instaurata fara ajutorul vreunei mîini\" si va acoperi întreg pamîntul (Dan. 2:34, 35, 45).\r\n\r\nCei dintîi crestini au cunoscut si ei aceste profetii din cartea lui Daniel si le-au dat o interpretare foarte precisa. Iata ce scria Hippolytus, care a trait între anii 160-236 dupa Cristos si a fost unul dintre ucenicii lui Irineu, socotit la rîndul sau ca unul dintre cei patru mari teologi ai vremii sale:\r\n\r\n„Capul de aur al Chipului, ca si leul corespunzator arata Imperiul Babilonian; pieptul si bratele de argint împreuna cu simbolul ursului din cap.7 sînt Imperiul Mezilor si Persilor; pîntecele si coapsele de arama împreuna cu pardosul (leopardul) îi arata pe Greci, care au detinut suprematia începînd cu vremea lui Alexandru Macedon; picioarele de fier ca si fiara „nespus de grozav de înspaimîntatoare si de puternica\" îi arata pe Romani, care sînt în fruntea lumii în vremea de azi; picioarele parte de fier si parte de lut, ca si cele zece coarne, sînt simboluri pentru zece împaratii care nu s-au ridicat înca; celalalt corn mai mic care s-a ridicat dintre cele zece îl arata pe Anticrist; piatra care loveste Chipul si aduce Judecata asupra întregului pamînt este Cristos\". („Tretise on Crist and Antichrist\" - Ante-Nicene Faters, Volumul V, pag.210, par.28)\r\n\r\nPartea de profetie care se împlinise deja le era foarte clara primilor crestini. Ei si-au dat seama chiar si ca în perimetru Imperiului Roman se vor mai defini înca alte zece forme de guvernamînt, aliate într-o forma de stapînire comuna, care îi va face loc Anticristului si ca Cristos, la cea de a doua venire a Sa, va pune capat „vremii neamurilor\", instaurînd chiar aici pe pamînt împaratia neprihanirii. Cei dintîi crestini nadajduiau înca de pe atunci în venirea împaratiei. Nu este de mirare, caci însusi Domnul îi învatase sa se roage spunînd: „Vie împaratia Ta\" (Mat. 6:10).\r\n\r\nRevenind la secventele visului lui Nebucadnetar si la talmacirea pe care i-o oferea Daniel ajungem sa privim la o veritabila perspectiva a istoriei. Iat-o:\r\n\r\nImperiul Babilonian - a fost instalat la putere începînd cu anul 604 înainte de Cristos, prin venirea lui Nebucadnetar la putere, ca un urmas al lui Nebopolasar. Acest imperiu a fost supranumit: „Un Imperiu de Aur, într-o epoca de aur\". Realizarile militare, economice si edilitare din Imperiu au depasit fara termen de comparatie tot ceea ce omenirea vazuse pîna la vremea aceea. Ca oras, Babilonul era o minunatie a lumii. Zidurile lui erau întinse pe un perimetru de 90 de km cu fiecare latura lunga de 23 de km, înalt de 70-100 de metri si gros de 25 de metri. În interiorul orasului existau provizii suficiente pentru a supravietui unui eventual asediu pe o durata de 20 de ani. Rîul Eufrat fusese deviat printre doi pereti dubli ai orasului si asigura alimentarea cu apa necesara. Nebucadnetar adunase în Babilon cantitati uriase de aur: cladise temple aurite, statui de aur si el însusi domnea pe un urias tron turnat din aur. Babilonul devenise un simbol al puterii omenesti si al trairii fara Dumnezeu.\r\n\r\nCa simbol profetic întîlnit de multe ori în profetii, Babilonul reprezinta încercarea oamenilor de a se descurca fara Dumnezeu, sistemul din care este exclusa închinarea la Dumnezeu si în care omenirea este organizata într-o societate suficienta de sine. În judecata lui Dumnezeu, un astfel de sistem nu trebuie sa dainuiasca. Chiar daca el nu-si gaseste un rival pe pamînt de care sa se teama, Dumnezeu însusi se ridica împotriva lui si-l sorteste pieirii. Într-un text din Isaia (45:1) ni se spune cum Dumnezeu l-a ridicat pe Cir-persanul si cum l-a ajutat în chip providential sa distruga Babilonul, înlesnindu-i patrunderea în cetate prin niste „porti, care sa nu se mai închida\". Istoria ne spune ca asa a si fost. Acest Cirus-persanul a fost ajutat sa asedieze Babilonul de Dariu-medul, un unchi batrîn al lui Cir. Ei au planuit ca în ajunul unei mari sarbatori din Babilon, cînd populatia orasului se îmbata de bucurie, sa devieze într-o depresiune apele Eufratului si sa patrunda prin albia goala a rîului, pe sub zidul cetatii, între cele doua ziduri ale fortificatiei. Totul ar fi ramas însa zadarnic, daca „cineva\" n-ar fi uitat tocmai în seara aceea sa închida portile de fier dinspre interiorul cetatii. Dumnezeu hotarîse ca falnicul Babilon sa devina „ca Sodoma si Gomora\", un loc în care nu va mai fi niciodata popor, ci îl vor locui fiarele pustiei si-l vor bîntui stafiile (Isaia 13:19- 22). Despre caderea Babilonului (în 538 î.Cr.) mai vorbise si Ieremia (Ieremia 50:1-3, 8-9, 14-16, 22-25; 51:1-4, 56-57) cu aproape 80 de ani înainte sa se întîmple. Amanunte despre luarea Babilonului gasim si în cartea lui Daniel, în capitolul 5:1-31.\r\n\r\nUrmasul lui Nebucadnetar, numit de Daniel: Belsatar, n-a învatat nimic din experientele tatalui sau si a fost înlaturat de la tron de însusi Dumnezeu, care-i cîntarise în balanta dreptatii Sale si-l gasise usor (Dan. 5:22-28). Caderea Babilonului a devenit sinonima cu înfrîngerea celor ce se ridica împotriva lui Dumnezeu, formînd un sistem prosper pentru o vreme, mîndru si arogant în înfatisare, dar gol si ignorant în esenta. Cartea Apocalipsei ne prezinta caderea unui alt Babilon, Babilonul cei mare. Numirea aceasta este simbolica bineînteles si face aluzie la asemanarile care vor exista între societatea viitorului, în timpul lui Anticrist si Babilonul istoric. Cititi în acest sens Apocalipsa 18:1-24. Concluziile acestui fragment din planul profetic revelat sînt multe. Noi ne marginim sa enumeram doar cîteva:\r\n\r\n(1) Dumnezeu face ce vrea în istoria lumii (subliniata de trei ori în Dan. 4:17, 25, 32).\r\n(2) Dumnezeu poate îngadui pentru o vreme nebunia oamenilor.\r\n(3) Mîndria merge înaintea caderii.\r\n(4) Oamenii nu învata din experienta trecutului.\r\n(5) Bogatia si slava lumii sînt puse în cîntarul divin si fiecare îsi va primi rasplata.\r\n\r\nImperiul Medo-Persan - a luat fiinta în 538 î.Cr. si a durat 200 de ani, pîna cînd a aparut pe scena lumii Alexandru Macedon în 331 î.Cr. Despre felul în care a cazut Babilonul în mîinile medo-persilor gasim scris în Daniel cap.5 (în special v.30-31). Dupa cum se vede din citirea textului, caderea Babilonului a fost hotarîta de însusi Dumnezeu care o anuntase de altfel cu aproximativ 100 de ani înainte prin Isaia (13:17-18). Urmasul lui Nebucadnetar n-a priceput nimic din experienta marelui împarat (Daniel 5:18-22) si a trebuit sa fie scos din scena istoriei. În visul lui Nebucadnetar, imperiul Medo-persan corespunde „pieptului si bratelor de argint\" (2:32), iar în vedeniile lui Daniel, el este asemanat cu „un urs care statea într-o rîna si avea trei coaste în gura între dinti\" (7:5), dar si cu berbecele din Dan. 8:1-4, 20). Aceste descrieri s-au potrivit întocmai cu caracteristicile acestui imperiu în care au fost introduse sisteme de taxare draconice si care si-a întins cuceririle dincolo de Egipt si pîna la granitele Greciei.\r\n\r\nImportanta profetica a imperiului Medo-Persan. Fara îndoiala ca evenimentul cel mai extraordinar a fost împlinirea absolut exacta a profetiei facuta despre aparitia si rolul împaratului persan Cir. Cu aproximativ 100 de ani înainte, Dumnezeu anuntase prin Isaia si numele acestui împarat si lucrarea pe care o va împlini acesta: sa dea voie evreilor sa se întoarca în patria lor (Isaia 45:1-13). Pe cînd slujea la curtea împaratului Dariu, Daniel „a vazut din carti\" ca s-au împlinit datele anuntate de profetul Ieremia (25:11) despre robia evreilor si împreuna cu mai marii poporului iudeu s-au înfatisat înaintea împaratului purtînd în mîini sulul profetiei lui Isaia. Impresionat de cele citite, Cir s-a pus imediat pe lucru si Ezra ne povesteste despre hotarîrea luata de el (Ezra l: l -11).\r\n\r\nDupa 200 de ani de dominare, Imperiul Medo-Persan s-a prabusit încercînd sa cucereasca Grecia. În planul lui Dumnezeu, venise vremea lui Alexandru Macedon. În batalia de la Arabela (331 î.Cr.) desi coplesiti numeric, proportia a fost se pare de unul la douazeci în favoarea Medo-Persanilor, Grecii au obtinut o victorie zdrobitoare. Împaratul medo-persan a încercat zadarnic sa-si regrupeze trupele într-o retragere strategica. „Pardosul\" grec (leopardul este cea mai rapida dintre feline) nu le-a dat nici un timp de ragaz, ci i-a urmarit peste tot cu o iutime nemaiîntîlnita în istorie. Pas cu pas, imperiul medo-persan a disparut, lasîndu-i loc lui Alexandru sa se întinda „pîna la marginile pamîntului\" (Daniel 8:5).\r\n\r\nImperiul Grecesc - este asemanat în visul lui Nebucadnetar cu „pîntecele si coapsele de arama\" (Daniel 2:32). Imaginea a fost cum nu se poate mai nimerita, deoarece grecii au fost primii în istoria lumii care au purtat în afara hainelor obisnuite echipament de razboi facut din arama. Despre soldatii greci se spunea ca sînt „de arama\". Coiful, platosele, si scutul le-au dat grecilor avantaje nete si i-au ajutat sa înfrînga osti care-i depaseau cu mult din punct de vedere numeric. Comandantul lor, Alexandru, a fost un geniu militar. Nimeni nu i-a putut sta împotriva. El a cucerit tot ceea ce putea fi cucerit în vremea aceea, iar dupa aceea, istoricii spun, a izbucnit în plîns, pentru ca nu putea merge mai departe.\r\n\r\nSuccesul l-a gasit pe Alexandru Macedon foarte repede. La nici treizeci de ani avea deja o putere incomensurabila. El a trait însa ca un smintit în betii si în serbari în care era proclamat una cu zeii. De fapt, chiar si campania lui împotriva Indiei, s-a nascut tot din dorinta de a repeta ceea ce facuse în traditia Olimpului grecesc Bachus si Hercules. La aproximativ 32 de ani, Alexandru Macedon a fost atins de friguri si a murit în delir dupa 11 zile de chin. Era anul 323 î.Cr.\r\n\r\nInfluenta imperiului grecesc asupra Israelului a fost adînca si de durata. Filozofia greceasca n-a facut casa prea buna cu religia evreilor, dar ocupantii s-au impus prin forta armata si viclenia distractiilor. Rezultatul influentei grecesti a fost o slabire a moralei si o provocare la adresa atasamentului evreilor fata de Dumnezeu.\r\n\r\nÎmparatul grec Antioh Epifaniu a atins culmea represaliilor atunci cînd a intrat calare în Templu si a oferit pe altar carne de porc (animal necurat în religia evreilor). El a omorît preotii si a încercat o grecizare fortata a evreilor. Prin ceea ce a facut, Antioh a intrat în istorie ca un precursor al lui Anticrist. (La el face aluzie Domnul Isus în Matei 24:15). Obraznicia cuceritorilor greci a trezit însa spiritul nationalist în evrei si sub conducerea Macabeilor, ei s-au revoltat si si-au dobîndit libertatea. Aceasta scurta perioada de libertate s-a sfîrsit însa repede. Din cauza framîntariior interne si a luptelor fraticide, Israelul a slabit treptat si a cazut sub stapînirea masinii de razboi romane. Venise vremea ultimului imperiu din visul lui Nabucadnetar si din vedeniile lui Daniel.\r\n\r\nImperiul Roman. În vedeniile lui, Daniel nu a gasit nici o asemanare suficient de socanta pentru a ilustra ceva din caracterul acestei puteri mondiale. Profetul scrie: „Era o a patra fiara, nespus de grozava de înspaimîntatoare si puternica\" (Daniel 7:7). Nimic n-a putut sa stea împotriva acestei uriase forte de invazie romana care „avea niste dinti mari de fier, mînca, sfarîma, si calca în picioare ce mai ramînea\"\r\n\r\nRoma si-a început cuceririle în anul 241 î.Cr., odata cu invadarea Siciliei. Sub talpile Legiunilor ei au cazut apoi Europa, Asia Mica, Orientul Mijlociu si coasta de nord a Africii, astfel ca întreaga Mediterana devenise un lac al Imperiului...\r\n\r\nÎn timp ce alte imperii au durat o perioada masurata în zeci sau sute de ani, puterea Romei a instaurat un Imperiu care a dainuit mai bine de l.500 de ani. Chiar daca ramura de apus a Imperiului a cazut cam pe la anul 500 d.Cr., Imperiul Bizantin a continuat pîna în 1453 d.Cr. De departe, Imperiul roman a fost cel mai mare dintre imperiile lumii.\r\n\r\nDin cele patru puteri mondiale anuntate de Dumnezeu prin visul lui Nebucadnetar si prin vedeniile lui Daniel, Imperiul roman este important într-un mod aparte:\r\n\r\n(1) el va fi cea din urma forma de stapînire omeneasca asupra pamîntului si\r\n\r\n(2) el va fi imperiul în care va lovi Cristos la cea de a doua Sa venire. Asupra acestui imperiu se merita sa aruncam o privire mai atenta si sa cautam sa identificam fragmentele de informatii pe care ni le pune la dispozitie textul profetic.\r\n\r\nCaracteristicile profetice ale imperiului fiarei.\r\n\r\na) Este un imperiu segmentat în doua sectiuni distincte. Textul din Dan. 2:33 ne spune ca în prima lui parte de existenta („picioare\"), imperiul este unitar si are taria fierului, în timp ce în cea de a doua parte a existentei lui (de la fluierele picioarelor în jos) imperiul se înfatiseaza sub forma unei imposibile unitati între fierul si lutul amestecate împreuna. Într-adevar, imperiul fiarei a patra a debutat sub numele de Imperiu al Romei si a supus sub talpile invincibilelor ei legiuni toata suflarea lumii. Romanii au cucerit prin forta si au dominat prin teroare. Fara a fi vreodata cucerit (!!!), acest imperiu al Romei a cazut într-un fel de „lesin\", disparînd temporar de pe scena istoriei. El s-a dat la o parte ca sa faca loc pentru „vremea Bisericii\". (Noi traim acum în vremea sfîrsitului si putem vedea cum imperiul Romei se reface, ca forma si ca alcatuire, sub numele de „Confederatia Europeana\" sau „Piata Comuna\" sau „Comunitatea Europeana\". Exista astazi chiar si un „Parlament European\" ca un mugure al unei visate conduceri unice.) În cea de a doua sa faza, Imperiul Romei sau Imperiul Fiarei, nu va mai fi unitar, ci fragmentat în zece unitati distincte de guvernamînt simbolizate de cele zece coarne ale fiarei (Daniel 7:7-8; Apoc.13:l). Faptul ca imperiul fiarei a patra se reface, este semnul ca vremea Bisericii este pe sfîrsite si ca Dumnezeu este gata sa reia firul profetiilor lui Daniel. În ordine cronologica urmeaza ca, dupa plecarea Bisericii, în imperiu sa apara „Omul faradelegii, Fiul pierzarii, Nelegiuitul\" (2 Tesal. 2:3-12; 1 Ioan 2:18; Apoc. 13:3-8), micul corn din vedenia lui Daniel 7:8, Anticristul, \"domnul care va veni\" (Daniel 9:26). El va face semne mari si minuni, va intra în Templul din Ierusalim si se va da drept Dumnezeu, va face razboi cu sfintii si-i va birui, dar va fi nimicit de „piatra desprinsa fara ajutorul vreunei mîini omenesti\" (Daniel 2:34-35). Acest Anticrist va fi nimicit de suflarea Domnului Isus care-l va prapadi la aratarea venirii Sale (2 Tesaloniceni 3:8). Dumnezeu va da apoi împaratia în mîinile Fiului Omului care va veni pe norii cerului (Daniel 7:13-14; 1 Corint. 15:24-27; Ps. 2).\r\n\r\nb) Este un imperiu care va avea o putere înfricosator de mare. Grozavia acestui imperiu nu si-a gasit echivalent în nici o creatura de pe fata pamîntului. Textele profetice îl numesc pur si simplu „FIARA\". Impresia pe care a lasat-o asupra lui Daniel a fost una „nespus de grozava de înspamîntatoare si de puternica\". „Avea dinti mari de fier si mînca, sfarîma si calca în picioare ce mai ramînea; era cu totul deosebita de toate fiarele de mai înainte\" (Daniel 7:7) Astazi oamenii au descoperit energia atomica si armele moderne de distrugere îngrozesc pe toti locuitorii pamîntului. Puterea de dominare a Fiarei se va sprijini si pe amenintarea cu razboiul atomic.\r\n\r\nc) Este un imperiu în care totul si toti se vor afla sub controlul celui aflat la putere. Apocalipsa ne spune ca fara asentimentul Fiarei „nimeni nu va putea sa vînda sau sa cumpere\" (Apoc.13:15-18). Mult timp aceasta profetie a parut greu de crezut, dar acum, de cînd au aparut calculatoarele electronice, computerele si credit-cardurile cu coduri numerice, toata lumea stie ca acesta este drumul pe care se îndreapta societatea de mîine.\r\n\r\nd) Este un imperiu cu o închinare impusa. Forme de închinare impusa au mai existat si în alte timpuri, dar ceea ce se va întîmpla în timpul domniei Fiarei va întrece orice închipuire. O a doua personalitate, la fel de dezumanizata (si numita din acest motiv: „o a doua fiara\", va îndemna tot pamîntul sa cada în admiratia Fiarei celei mari:\r\n\r\n„Ea lucra cu toata puterea fiarei dinaintea ei (aceiasi putere satanica deci) si facea ca pamîntul si locuitorii lui sa se închine fiarei dintîi, a carei rana fusese vindecata...  Ea a zis locuitorilor pamîntului sa faca o icoana fiarei, care avea rana de sabie si traia. I s-a dat putere sa dea suflare icoanei fiarei, ca icoana fiarei sa vorbeasca si sa faca sa fie omorîti toti cei ce nu se vor închina icoanei fiarei\" (Apocalipsa 13:12-15).\r\n\r\ne) este primul si singurul imperiu care, crezînd în Dumnezeu, se va ridica declarat împotriva Lui. Forme de împotrivire fata de Dumnezeu au existat întodeauna, dar de obicei oamenii aceia erau sau agnostici sau atei sau idolatri. Fiara care va conduce imperiul roman restaurat, va fi pe fata Anti-Dumnezeu, Anti-Cristos. Diavolul însusi o va însufleti dupa încercarea de atentat sugerata de „rana de moarte care fusese vindecata\" (Apoc. 13:12) si îi va transmite împotrivirea lui fata de Stapînul universului care l-a facut „sa-si piarda vrednicia\" si sa devina „Satanah\" (Împotrivitorul).\r\n\r\nFiara „va intra în Templul lui Dumnezeu, dîndu-se drept Dumnezeu\" (2 Tesal. 2:4).\r\n\r\n„El va rosti vorbe de hula împotriva Celui Prea înalt, si se va încumeta sa schimbe vremile si legea; si sfintiilor fi dati în mîinile lui timp de o vreme, doua vremi si jumatate de vreme\" (Daniel 7:25).\r\n\r\n„Împaratul va face ce va dori; se va înalta mai pe sus de toti dumnezeii si va spune lucruri nemaiauzite împotriva Dumnezeului dumnezeilor; si va propasi pîna va trece mînia, caci ce este hotarît se va împlini. Nu va tine seama nici de dumnezeii parintilor sai, nici de dorinta femeilor; cu un cuvînt, nu va tinea seama de nici un dumnezeu, ci se va slavi pe sine mai pe sus de toti\". (Daniel 11:36-37)\r\n\r\nDespre acest imperiu al Anticristului gasim o relatare mult mai amanuntita în cartea Apocalipsei. Pentru studiul de fata este bine sa ne oprim aici si sa ne amintim ca aceasta ultima forma de guvernamînt al oamenilor, mai rea si mai perversa decît toate celelalte, va fi pedepsita nu prin interventia vreunei alte structuri sociale, ci prin însasi interventia lui Dumnezeu, care în persoana lui Cristos („piatra desprinsa fara ajutorul vreunei mîini omenesti\"- Daniel 2:34) se va napusti asupra chipului vazut de Nebucadnetar, punînd capat vremii neamurilor, si întemeind o noua împaratie, care va fi data „poporului sfintilor Celui Prea înalt\" (Daniel 7:27).\r\n\r\n„M-am uitat în timpul vedeniilor mele de noapte si iata ca pe norii cerului a venit unul ca un fiu al omului; a înaintat spre Cel îmbatrînit de zile si a fost adus înaintea Lui. I S-a dat stapînire, slava si putere împarateasca, pentru ca sa-i slujeasca toate popoarele, neamurile, si oameni de toate limbile. Stapînirea Lui este o stapînire vesnica, si nu va trece nicidecum, si împaratia Lui nu va fi nimicita niciodata\" (Daniel 17:13-14).\r\n\r\nFara sa staruim prea mult la analizarea visului lui Nebucadnetar, este bine totusi sa subliniem doua concluzii care ne afecteaza pe noi, cei de astazi:\r\n\r\n(1). Sfîrsitul epocii noastre nu se va produce în urma unei îmbunatatiri treptate care sa culmineze cu venirea împaratiei cerurilor, ci se va împlini printr-un moment de criza, de prabusire si de catastrofa neasteptata. În vedenia lui Nebucadnetar, peste degetele de fier si lut se prabuseste „piatra, deslipita fara ajutorul vreunei mîini\" (aratata în cap.7 a fi Cristos în împaratia lui Mesianica) si face bucati întreg Chipul, transformîndu-l într-o pleava luata de vînt si împrastiata fara urma (cap.2:34-35, 43-45).\r\n\r\nIata ce scrie William Newall în aceasta privinta: „Toate visarile moderne despre un Mileniu înainte de venirea lui Cristos sînt erezii nascute dintr-o necugetata încredere în bunatatea oamenilor sau, si mai grav, din înselare Satanica. „Cînd vor zice: „Pace si liniste!\" atunci o prapadenie neasteptata va veni peste ei, ca durerile nasterii peste femeia însarcinata, si nu va fi chip de scapare\" (1 Tesal.5:3). Cristos va face razboi cu lumea actuala si va spulbera puterile lumii transformîndu-le în pleava purtata de vînt\".\r\n\r\n(2). Sfîrsitul vremurilor noastre este acum aproape. Cele doua picioare ale Chipului din visul lui Nebucadnetar sînt o reprezentare fidela a istoriei. Dupa cîte stim, imperiul roman s-a rupt în doua ramuri - Imperiul de Rasarit si Imperiul de Apus. Despartirea s-a produs în anul 395 d.Cr. Fara sa riscam prea mult, noi credem ca ne aflam astazi în perioada descrisa de imaginea celor zece degete ale picioarelor Chipului din vis. Traim astazi o renastere a Imperiului roman pe teritoriul cunoscut din Europa. Slabiciunea fierului amestecat cu lutul, defineste cum nu se poate mai bine, lipsa de coeziune care exista sub masca unitatii europene. „Legaturile omenesti\" de care vorbea Daniel exista: casatorii peste granite, calatorii fara pasapoarte, investitii multinationale, libertatea companilor de a prelua contracte în alte tari, planul de a avea în curînd o moneda comuna unica, etc. Totusi, dominarea Europei de astazi asupra lumii nu va mai avea forta si taria Imperiului de altadata. Dominarea militara va fi înlocuita cu dominarea economica si politica. ªubrezenia partilor componente îi va determina pe cei din noua Europa Unita sa-l accepte ca unic conducator pe Anticrist. Cu solutiile propuse de el, lumea va parea ca iese din impas, dar... pacea adevarata nu va fi asezata decît de venirea Domnului Pacii. Va mai fi un groaznic razboi mondial: Armaghedonul, sau „ultimul razboi dinaintea pacii\".\r\n\r\nCELE 70 DE SaPTaMÎNI\r\n\r\nVedenia din cel de al 9-lea capitol al cartii lui Daniel este o profetie privitoare la Israel si la evenimentele cheie prin care va trece acest popor „de la darea poruncii pentru zidirea din nou a Ierusalimului, pîna la Unsul (Mesia), Cîrmuitorul\" (Dan. 9:25).\r\n\r\nAceasta perioada cuprinde istoria din „vremea neamurilor\", pîna cînd Dumnezeu se va arata si va reaseza Israelul în prerogativele Mesianice („toate neamurile vor fi binecuvîntate în samînta ta\" - Gen 22:18).\r\n\r\nIata cuprinsul profetiei:.\r\n\r\n(a) În versetele 24-27 i se spune lui Daniel ca: „70 de saptamîni au fost hotarîte asupra poporului sau\". Aceste 70 de saptamîni (sau septade, deoarece în limba folosita grupa de 7 nu are neaparat semnificatia de saptamîna) sînt împartite în doua grupe, dupa cum urmeaza: primele 69 de saptamîni si saptamîna a-70-a.\r\n\r\nDe la darea poruncii pentru rezidirea Ierusalimului\" si pîna în ziua cînd Mesia va fi „stîrpit” au fost hotarîte sapte saptamîni si sase zeci si doua de saptamîni\". Facînd socotelile ajungem la 69x7=483 de ani de la decret si pîna la moartea lui Mesia. Ramîne înca o a 70-a saptamîna de ani, rezervati pentru vremea sfîrsitului, cînd un „domn al unui popor care va veni\" va face un legamînt trainic cu Israelul, dar la mijlocul acestei saptamîni va calca legamîntul si, intrînd în Templu, se va da drept Dumnezeu.\r\n\r\nPentru a întelege profetia trebuie sa stabilim data la care s-a dat decretul pentru rezidirea Ierusalimului. Trecînd peste cele trei decrete amintite de Ezra în cartea sa si care au lasat Ierusalimul cu zidurile înca neridicate, ajungem la porunca data de Artaxerxe ca urmare a cerem facute de Neemia: „trimite-ma în Iuda, la cetatea mormintelor parintilor mei, CA S-O ZIDESC DIN NOU\" (Neemia 2:5). Data acestui decret ne este data la începutul cartii: „În luna Nisan a anului al douazecelea al împaratului Artarxerxe\" sau în luna Nisan 445 î.Cr.\r\n\r\nDupa obiceiul evreiesc, atunci cînd nu ni se precizeaza ziua din luna este vorba de cea dintîi zi a ei. Facînd transformarile corespunzatoare calendarului Iulian, Sir Robert Anderson în colaborare cu „Astronomer Royal\" au ajuns la concluzia ca aceasta data a fost 14 Martie 445 î.Cr. În socotelile profetice vorbeste în anii luni-solari de cîte 360 de zile (vezi toate socotelile lui Ioan în cartea Apocalipsei) si daca facem socotelile de rigoare ajungem exact la ziua în care a intrat Domnul Isus în Ierusalim. Nu se poate sa nu fii socat de precizia cu care s-au împlinit prezicerile lui Daniel. Cu pasiunea lui pentru amanunte, Evanghelistul Luca ne spune ca aceste lucruri se întîmplau „în cel de al cincisprezecelea an al domniei lui Tiberiu Cezar\" (Luca 3:1). ªtiind ca Tiberiu si-a început domnia la 19 August anul 14 d.Cr. ajungem cu socoteala la anul 29 d.Cr. sau exact 483 de ani de la porunca pentru rezidirea Ierusalimului.\r\n\r\nCine a citit Evangheliile stie ca Domnul Isus s-a retras într-un fel de obscuritate deliberata dupa ce a fost refuzat de mai marii evreilor asteptînd ceea ce El numea „Ceasul\" care-i fusese hotarît. Era vorba de ceasul profetic al împlinirilor Mesianice (Ioan 2:4; 7:8; 8:20).\r\n\r\n„Acum sufletul Meu este tulburat. ªi ce voi zice?... Tata, izbaveste-Ma din CEASUL acesta?... Dar tocmai pentru aceasta am venit pîna la CEASUL acesta!\" (Ioan 12:27).\r\n\r\nÎntr-adevar în repetate rînduri El le-a poruncit ucenicilor sa nu spuna nimanui despre divinitatea Sa, dar în ziua Floriilor a primit slava multimilor care L-au primit ca pe Mesia! Cînd fariseii au protestat, invocînd blasfemia si teama de represaliile romanilor, Domnul Isus le-a spus clar: „Aceasta este ziua! Va spun ca daca vor tacea ei, pietrele vor striga!\" (Luca 19:40)\r\n\r\nFara a-si face iluzii, El a privit de departe Ierusalimul si a izbucnit în plîns spunînd: „Daca ai fi cunoscut si tu macar în aceasta zi, lucrurile care puteau sa-ti dea pacea! Dar acum ele sînt ascunse de ochii tai. Vor veni peste tine zile cînd vrasmasii tai te vor înconjura cu santuri, te vor împresura si te vor strînge din toate partile te vor face una cu pamîntul, pe tine si pe copiii tai din mijlocul tau; si nu vor lasa în tine piatra peste piatra, PENTRU Ca N-AI CUNOSCUT VREMEA CÎND AI FOST CERCETATA\" (Luca 19:42-44)\r\n\r\nIerusalimul ar fi trebuit sa stie ca se împlinisera cei 483 de ani si ca venise vremea profetiei rostita de Zaharia: Împaratul, Cîrmuitorul venise blînd si calarepe un mînz, pe mînzul unei magarite (Zaharia 9:9; Matei 21:4-5). Se împlnisera cei 483 de ani de asteptare. „Unsul\" înainta spre criza suprema a calatoriei lui terestre. Se apropia de momentul cînd avea sa fie „stîrpit\", si nimeni în afara de El nu stia înca ce vroise sa spuna Daniel atunci cînd adaugase „stîrpit, si nu va avea nimic\" (9:26). Mîntuitorul calatorea înspre înviere!\r\n\r\nCe se întîmpla însa cu saptamîna a-70-a? Ea ramîne înca în viitor. Între „stîrpirea” Unsului la sfîrsitul celei de a 69 saptamîni si începutul celei de a 70-a saptamîni se întinde acum „taina tinuta ascunsa de veacuri\": vremea Bisericii. Biserica a fost un secret pe care Dumnezeu nu l-a descoperit celor din Israel (Efeseni 3:2-13). La sfîrsitul acestei perioade, în care harul este „la neamuri\", Dumnezeu se va întoarce iarasi înspre Israel, îi va readuce în patria lor si vor trai împreuna cu toata omenirea vremea amagitoare a Anticristului. Sub protectia lui initiala, ei îsi vor putea reconstrui Templul, dar la jumatatea saptamîna, legamîntul va fi rupt, Anticristul va intra în Templu dîndu-se drept Dumnezeu si declansmd lantul violent de evenimente relatat pe larg în cartea Apocalipsei.\r\n\r\nDin pacate nu putem merge mai departe în analizarea profetiilor lui Daniel. Speram însa ca acest studiu asupra celor doua pasaje profetice de baza sa va fie folositor pentru investigatii personale ulterioare. Între timp, cu cea de a-70-a saptamîna în minte, noi asteptam sunetul trîmbitei divine, vocea arhanghelul, coborîrea Domnului, deschiderea mormintelor, învierea celor sfinti, stapînirea din timpul împaratiei, si slava de care vor fi urmate toate acestea.\r\nCauta in biblie\r\nTextul cautat\r\nCauta...\r\nCum caut\r\nToate Cuvintele\r\nUnde caut:\r\nToată biblia\r\nVersetul Zilei\r\nJudecatori 6:14-16\r\n\r\nDu-te cu puterea aceasta pe care o ai, si izbaveste pe Israel din mana lui Madian! oare nu te trimit Eu\r\nMeditatia Zilei\r\n\r\nPowered by Biblia Online\r\nFacebook\r\n\r\nParteneri\r\nStiri Crestine\r\nMuzica Crestina\r\nVersuri Crestine\r\nChristian Lyrics\r\nBiblia Online Cornilescu\r\nVersuri Crestine\r\nNu este Crăciun fără Isus - Estera & Laura Bretan\r\nÎn iesle azi - Diana Scridon Pop\r\nBucurie in suflet - Familia Timofte\r\nSe-aude Glas Peste Carpați / Tăria Noastră Fii Mereu - The Beuca Family\r\nUn singur Dumnezeu - Otto Pascal & Biji\r\nFiu iubit - BBSO\r\nClipa de clipa - Philadelphia Band\r\nPrintul Pacii a venit - Alin și Emima Timofte & TB Music Family\r\nPraise - Lumina Worship\r\nUnde ești țara mea?\r\nPowered by VersuriCrestine.ro\r\n" },
    //OSEA
    { "OSEA", "Explicatia Cartii: Osea\r\n\r\nCu Osea intram în „cartile celor 12\". Acest grup de profeti mai sînt numiti si „profetii mici\" din cauza continutului redus al profetiilor lor. În nici un fel însa, ei nu sînt „mici\" ca însemnatate. Scurte ca niste telegrame duhovnicesti, cartile lor sînt la fel de urgente si de convingatoare. Iata-l de exemplu pe Osea: cartea lui vorbeste despre infidelitatea fata de Dumnezeu, prezentînd-o ca pe un adulter spiritual. În ultima analiza, orice pacat este o teribila infidelitate în dragostea noastra fata de Domnul. Pacatul îl raneste pe Dumnezeu si ne distruge pe noi însine. Cartea poarta numele profetului care a scris-o. Osea. Iosua si Isus sînt forme derivate din aceiasi radacina. „Osea\" înseamna „mîntuire\", în timp ce Iosua si Isus includ în plus prezenta celui care produce mîntuirea: „Mîntuirea este a Domnului\" sau „Domnul este Mîntuitorul\".\r\nOsea n-a fost un profet crescut în scoala profetilor, ci un om ridicat de Dumnezeu din mijlocul poporului. El a trait în împaratia de nord a Israelului. Tatal sau s-a numit Beeri (Osea 1:1), iar sotia lui Gomer (Osea 1:3). Din casnicia lor s-au nascut trei copii: doi baieti si o fata (Osea 1:4, 6, 9) care au servit ca semne cu un mesaj simbolic pentru Israel. Alte amanunte nu mai cunoastem despre Osea. Numele lui nu mai apare în nici o alta carte a Bibliei.\r\nOsea si-a rostit mesajul pentru poporul împaratiei lui Israel (Osea 5:1), pe vremea lui Ozia (767-739 î.Cr.), Iotam (739-731 î.Cr.), Ahaz (731-715 î.Cr.) si Ezechia (715-686 î.Cr.), toti acestia fiind din împaratia lui Iuda. Probabil ca dupa caderea Samariei si ducerea lui Israel în robie, Osea s-a refugiat în împaratia lui Iuda si de aceea îsi dateaza activitatea dupa împaratii care erau familiari noilor sai cititori. Cînd si-a început Osea activitatea în Israel, Ieroboam al doilea mai era înca pe tron (782-753 î.Cr.). Aceasta îl face pe Osea un contemporan mai tînar al profetului Amos. De fapt, activitatea lui a fost paralela cu activitatile profetilor Isaia si Mica, care si-au rostit mesajele în împaratia lui Iuda. Îndelunga activitate profetica a lui Osea a acoperit aproape 50 de ani, strabatînd între 755-710 î.Cr. vremea ultimilor 6 împarati dinaintea ducerii Israelului în robia Asiriana.\r\nPerioada de timp dintre domnia lui Ieroboam si robia babiloniana este „ultima turnanta\" din traseul istoric al împaratiei lui Israel. Din ce în ce mai mult, lucrurile au început sa se precipite înspre rau. Ieroboam al doilea a fost ultimul împarat de pe tronul Israelului care a fost instalat prin ceva ce a semanat macar cu o alegere divina. Cu moartea sa si cu asasinarea fiului sau Zaharia (2 Împ. 10:30 si 15:8-12), ia sfîrsit dinastia instituita de Iehu. Cei care se vor succeda de acum la tron, vor pune mîna pe putere prin lovituri de stat. ªalum îl asasineaza pe Zaharia dupa numai o luna de domnie; Menahem îl asasineaza pe ªalum, tot dupa o singura luna petrecuta pe tron. Pecah îl asasineaza si el pe Pecahia, iar Osea (împaratul, nu profetul) ajunge la tron prin asasinarea ucigasului Pecah.\r\n\r\nA fost o perioada cumplita de istorie. Loialitatea fata de tron era ca si inexistenta. În toate ungherele se ascundeau si complotau conspiratorii. Anarhia se ridica valuri, valuri, iar poporul pierduse orice stabilitate si orice masura (Osea 4:1, 2; 7:1, 7; 8:4; 9:15). În jurul tronului pîngarit si însîngerat, natiunea putrezea în imoralitate si idolatrie. Instabilitatea politica a slabit mult taria Israelului. Încrederea lor în Domnul era inexistenta! tara alerga între Egipt si Asiria dupa aliante internationale (Osea 7:11). Din punct de vedere spiritual, situatia era si mai grea de cît situatia politica. Imediat dupa scindarea de Ierusalim si de casa lui David, Ieroboam facuse Israelul sa pacatuiasca cu viteii de aur pe care-i instalase la Dan si Betel. Fusese o vreme în care idolatria îl simboliza totusi macar pe Iehova, Dumnezeul evreilor (1 Împarati 12:25-33). Cu timpul însa, idolul a capatat putere prin sine însusi si Israelul l-a parasit pe Domnul. Viteii de aur au devenit în scurt timp o usa deschisa pentru idolatria cea mai rudimentara si mai îndracita. Astarteele si Baali au intrat în inimile poporului, întunecîndu-le mintile si tîrîndu-i în cele mai murdare ritualuri si în crudele sacrificii ale copiilor. Iata o lista incompleta a relelor împotriva carora a vorbit Osea:\r\n\r\n- necinste (4:1, 2),\r\n- crime si varsari de sînge (4:2; 5:2; 6:8),\r\n- hotii la drumul mare savîrsite si de tîlhari, dar si de preoti (6:9; 7:1),\r\n- imoralitate practicata pe scara nationala (4:2, 11; 7:4),\r\n- necinste în comert si în justitie (10:4; 12:7),\r\n- idolatrie demonica (4:12-13; 8:5; 10:1, 5; si extraordinarul 13:2: „ªi jertfind oameni, saruta vitei!\"),\r\n- betie (4:2; 7:5),\r\n- o totala lipsa de sensibilitate si de cainta fata de Dumnezeu (4:4; 13:14).\r\n\r\nTrist tablou al unui popor care facuse odinioara un legamînt de credinciosie fata de Dumnezeu si fata de Legea Sa!\r\nÎn cei 50 de ani de misiune profetica, Osea si-a repetat mereu cele trei parti ale mesajului sau:\r\n\r\n(1) Dumnezeu este mîniat la culme pentru pacatele în care traieste poporul Sau si din pricina aceasta\r\n(2) judecata este facuta si pedeapsa este sigura; totusi\r\n(3) dincolo de aceasta mînioasa pedepsire, Dumnezeu pastreaza poporului o iubire vesnica, plina de bunatate si credinciosie, care pregateste deja planuri pentru timpul în care ei se vor întoarce la El.\r\n\r\nPrimele 3 capitole ale cartii sînt autobiografice si simbolice. Osea este îndemnat de Dumnezeu sa se casatoareasca cu o femeie pe nume Gomer. Casnicia profetului cu aceasta femeie necredincioasa si dedata la pacatul curviei, devine o alegorie a tragediei pe care o traieste Dumnezeu în relatia Lui cu nestatornicul si necredinciosul Israel. Este evident ca, în succesiune logica, capitolul 2 este o „talmacire\" spirituala a situatiei neplacute în care l-a pus Dumnezeu pe Osea. Urmeaza apoi capitolul 3 al cartii, care priveste adînc în viitor pîna spre vremea sfîrsitului în care Israelul se va întoarce la Domnul: „Dupa aceea, copiii lui Israel se vor întoarce si vor cauta pe Domnul, Dumnezeul lor, si pe împaratul lor David; si vor tresari la vederea Domnului si a bunatatii Lui, în vremurile de pe urma\" (Osea 3:5).\r\n\r\nRestul de capitole ale cartii sînt o culegere de pasaje retorice ale profetului. Este foarte greu sa stabilesti o ordine oarecare în aceste fragmente de cuvîntari înflacarate. Ele sînt izbucniri de gelozie din partea unui Dumnezeu al iubirii.\r\n\r\nCuvinte cheie si teme caracteristice: Osea este cunoscut prin faptul ca prezinta lipsa de credinciosia a poporului Israel fata de Dumnezeu drept o infidelitate spirituala. Cuvintele rostite de el sînt tari si taioase. Idolatria este numita „curvie\", iar idolii vremelnici sînt numiti „ibovnici\" (Osea 2:2-5).\r\n\r\nÎn noianul de cuvinte tari exista totusi si un cuvînt duios si plin de dragoste. El este „hesed\" si este unul dintre cele mai perfecte echivalente pentru dragostea dumnezeiasca. „Hesed\" este iubirea compatimitoare si dezinteresata. Hesed este dragostea statornica si imposibil de înlaturat. David a folosit cuvîntul acesta, atunci cînd a cautat pe cineva din casa vrasmasa a fostului împarat Saul pentru ca dorea sa le faca un bine: „David a zis: „A mai ramas cineva din casa lui Saul, ca sa-i fac bine din pricina lui Ionatan?\" (2 Samuel 9:1). tineti minte ca David facuse un legamînt cu Ionatan, prin care se angajasera sa-si faca bine unul altuia (1 Sam. 18:3; 20:14-16; 42). Pasajele în care Osea vorbeste despre „hesed\" sînt: Osea 2:19; 4:1; 6:4, 6; 10:12; 12:6.\r\n\r\nIata cum este descrisa manifestarea lui „hesed\" în cadrul relatiei pe care o are Dumnezeu cu Israelul:\r\n\r\n„Poporul Meu este pornit sa se departeze de Mine; si daca sînt chemati înapoi la Cel Prea înalt, niciunul dintre ei nu cauta sa se ridice. „Cum sa te dau Efraime? Cum sa te predau Israele? Cum sa-ti fac ca Admei? Cum sa te fac ca teboimul? Mi se zbate inima în Mine, si tot launtrul Mi se misca de mila! Nu voi lucra dupa mînia Mea aprinsa, nu voi mai nimici pe Efraim; caci Eu sînt Dumnezeu, nu un om. Eu sînt Sfîntul în mijlocul tau, si nu voi veni sa prapadesc\" (Osea 11:7-9).\r\n\r\nHesed este dragostea în virtutea Legamîntului încheiat între Dumnezeu si oameni.\r\nCartea lui Osea este o chemare la pocainta. Capitolul 14 al cartii este un mesaj de dragoste si de chemare pentru toti aceia care s-au îndepartat vreodata de Domnul. Osea mai este însa si altceva. Profetia aceasta este un avertisment împotriva idolatriei si a îndepartarii de Dumnezeu. Chiar si formele simbolice care pretind ca-L arata pe Domnul nu sînt altceva decît o cursa pentru suflet. Toate apostaziile mari au aparut la început ca o mica si neesentiala departare de la simplitatea si claritatea mesajului Scripturii. Din acest punct de vedere, am putea spune ca Ieroboam al doilea si ceilalti împarati de dupa el, au cules doar ceea ce semanase la începutul istoriei lui Israel, primul Ieroboam, fauritorul celor doi vitei de aur. Am putea reformula avertismentul lui Osea în cuvintele apostolului Ioan: „Copilasilor, feriti-va de idoli\" (1 Ioan 5:21).\r\nIntroducere - O patanie cu tîlc, 1 - 3\r\n\r\nI. Dumnezeu este sfînt 4-7\r\nCele cinci învinuiri, 4-5\r\nFalsa „întoarcere\", 6\r\nO vindecare imposibila, 7\r\n\r\nII. Dumnezeu este drept, 8-10\r\nTrîmbita anunta pedeapsa, 8:1-etc.\r\nTextul este o vestire a pedepsei care va veni.\r\n\r\nIII. Dumnezeu este dragoste 9-14\r\nUn Dumnezeu care tînjeste (11:1, 4, 8, etc.)\r\nUn Israel care nu poate scapa nepedepsit (12)\r\nTriumful final al iubirii (14)" },
    //IOEL
    { "IOEL", "Explicatia Cartii: Amos\r\n\r\nIata o carte scrisa de un profet laic. Dumnezeu si-a pastrat întotdeauna dreptul sa se foloseasca de orice oameni pe care El îi alege pentru vestirea mesajelor Sale. Cînd preotii si împaratii îsi pierd vrednicia, ciobanii si culegatorii de smochine sînt înrolati în vestirea Cuvântului Domnului (Amos 7:14). Cartea poarta numele autorului ei: Amos, care se poate traduce prin „purtatorul de poveri\". Înca o data, numele este o anuntare a specificului misiunii pe care o va avea acest om chemat de Dumnezeu la o activitate grea si importanta.\r\nNumele lui Amos nu apare în nici o alta carte a Bibliei. El nu a fost un preot sau un profet educat în mînuirea si transmiterea Cuvîntului lui Dumnezeu. Ca si în cazul celorlalti profeti însa, autoritatea lui a venit de la Domnul si asta i-a fost de ajuns ca sa stea plin de îndrazneala în fata oamenilor:\r\n\r\n„Amos a raspuns lui Amatia: „Eu nu sînt nici prooroc, nici fiu de prooroc; ci sînt pastor, si strîngator de smochine de Egipt. Dar Domnul m-a luat de la oi, si Domnul mi-a zis: „Du-te si prooroceste poporului Meu Israel!\" (Amos 7:14-15).\r\n\r\nLocul din care s-a ridicat Amos a fost Tecoa, o asezare situata cam la 6 km sud de Betleem, în pustia iudeii. Acolo îsi pastorise altadata si David oile si tot acolo se ascunsese pe cînd fugea de împaratul Saul. Limbajul folosit de Amos este nesofisticat. Profetul are vorbirea unui om de la tara, plina de comparatii din natura în care-si duce viata (Amos 3:4, 5, 12; 5:8, 19; 9:9). Tonul lui este raspicat, iar mesajul lui merge direct la tinta fara prea multa pregatire: „El a zis: „Domnul racneste din Sion, glasul Lui rasuna din Ierusalim. Pasunele pastorilor jalesc, si vîrful Carmelului este uscat\" (Amos 1:2).\r\nAutorul fixeaza singur data activitatii sale: „pe vremea lui Ozia, împaratul lui Iuda si pe vremea lui Ieroboam, fiul lui Ioas, împaratul lui Israel, cu doi ani înaintea cutremurului de pamînt\". tinînd seama de amenintarile care ne apropie de sfîrsitul domniei lui Ieroboam, putem spune ca Amos a activat în preajma anului 755 î.Cr.\r\n\r\nAmos este un profet ridicat de Dumnezeu din împaratia lui Iuda si trimis sa prooroceasca la Betel si Samaria, în inima religioasa a împaratiei lui Israel. Activitatea profetica a lui Amos nu a fost solitara. Probabil ca în tineretea lui a mai apucat sa-l vada pe Elisei si sa-l auda pe Iona vorbind despre succesele militare ale Israelului (2 Împarati 14:25) ). Amos a fost premergatorul altor profeti. În timpul lucrarii lui, Dumnezeu l-a ridicat pe Osea sa vorbeasca în Israel si pe Isaia si Mica sa activeze în Iuda.\r\nAmos a profetit într-o vreme de optimism si abundenta. Vremea lui Ozia si Ieroboam readusese ceva din fala Israelului de altadata. La curtile împaratesti se instalase iarasi luxul, ritualurile religioase erau pline de pompa si fast, poporul traia iarasi în abundenta si în bucurie. Israelul arata însa ca un mar frumos pe dinafara si gaunos pe dinauntru. Pacatul se furisase în viata intima a poporului si dincolo de aparentele de religiozitate, Legea Domnului fusese abandonata. Simplitatea si neprihanirea fusesera înlocuite cu abundenta materiala, luxul, necinstea si imoralitatea (Amos 2:6-8; 3:10; 4:1; 5:10-12; 8:4-6). În plan international, Asiria, Mesopotamia si Egiptul erau în eclipsa, asa ca aparent, Israelul nu avea de ce sa se teama. În acest context Amos a fost trimis de Domnul sa predice un mesaj de mustrare si amenintare. Pedeapsa vestita de Amos nu se vedea atunci nici macar la orizontul istoriei, asa ca nu este de mirare ca acest „taran\" venit de la oi a fost privit cu curiozitate si cu aversiune în centrele civice si religioase ale Israelului. Imaginati-va cam cum au reactionat doamnele din „înalta societate\" a Samariei cînd l-au auzit spunîndu-le: „Ascultati acum, juncane din Basan, de pe muntele Samariei, voi care asupriti pe cei sarmani, zdrobiti pe cei lipsiti, si ziceti barbatilor vostri: „Dati-ne sa bem!\" (Amos 4:1). Nu este de mirare ca împotriva profetului s-a pornit repede o conspiratie care a facut totul ca sa-l reduca la tacere. Amatia, preotul idolatru din Betel, l-a parît împaratului, declarîndu-l tradator (Amos 7:10-11), iar apoi l-a interpelat public, gonindu-l din Israel:\r\n\r\n„Pleaca, vazatorule si fugi în tara lui Iuda! Manînca-ti pîinea acolo, si acolo prooroceste. Dar nu mai prooroci la Betel, caci este un locas sfînt al împaratului, si este un templu al împaratiei\" (Amos 7:12-13).\r\n\r\nAceasta interpelare a cauzat, declaratia autobiografica a profetului si groaznica sentinta dumnezeiasca rostita împotriva preotului Amatia (Amos 7:17). Alungat din Israel, Amos s-a întors în Iuda si a asternut continutul profetiilor sale în scris ca sa poata fi raspîndite în continuare în Israel si ca sa serveasca drept avertisment împaratiei lui Iuda.\r\nDesi de la tara, Amos este un vorbitor elocvent si ordonat. Capitolele 1 si 2 sînt o expunere despre „opt greutati\" care stau pe inima profetului. În revarsari profetice, Amos vorbeste împotriva tuturor celor opt natiuni din teritoriile Palestinei: Siria (1:3-5), Filistenii (1:6-8), Fenicia (1:9, 10), Edom (1:11-12), Amon (1:13-15). Moab (2:1-3), Iuda (2:4-5) si Israel (2:6-16). Urmeaza apoi capitolele 4, 5 si 6 care contin rezumatele a trei predici rostite de Amos. Începutul acestor trei mesaje este usor de gasit. Fiecare dintre ele încep cu: „Ascultati cuvîntul acesta...” (3:1; 4:1; 5:1). Ultimele capitole ale cartii cuprind o serie de 5 viziuni prin care, într-o succesiva accelerare a mîniei este vestita pedeapsa care va veni asupra celor ce s-au îndepartat de Domaul (Amos 7-9).\r\n\r\nCuvinte cheie si teme caracteristice: Doua expresii au pasit afara din cadrul cartii lui Amos înspre frazeologia tipologica crestina:\r\n\r\n„Pregateste-te sa te întîlnesti cu Dumnezeu tau\" (Amos 4:12) a devenit semnal de atentionare valabil oriunde si oricînd.\r\n\r\n„Vai de cei ce traiesc fara grija în Sion\" (Amos 6:1) a devenit sinonim cu avertizarea celor care se afla în pericolul formalismului religios.\r\nCartea lui Amos este bogata în mesaje spirituale. Ea raspunde întrebarii: „Cine poate sa slujeasca Domnului?\" cu un raspuns care trece dincolo de preotia traditionala, la masa mare a poporului, stabilind drept unic criteriu de calificare „chemarea Domnului pentru slujba\". Amos denunta ipocrizia si falsul ritualului religios lipsit de substanta reala a unei vieti traite în ascultare de Cuvîntul Domnului: „Duceti-va numai la Betel, si pacatuiti!... Aduceti-va jertfele în fiecare dimineata, si zeciuielile la fiecare trei zile! Faceti sa fumege jertfe de multumire facute cu aluat! Trimbitati-va, vestiti-va darurile de mîncare de buna voie! Caci asa va place, copii ai lui Israel, zice Domnul Dumnezeu\" (Amos 4:4-5).\r\n\r\nCartea lui Amos ne vorbeste despre iluzoria siguranta a belsugului si confortului zilei de astazi. Dumnezeu este acela care ridica sau coboara împaratii si împaratiile. Toate marile imperii ale lumii au cazut din cauza aceluiasi motiv groaznic: pacatul. El atrage ruina popoarelor. Peste Israelul prosper rasuna glasul de tunet al lui Amos: „Asa vorbeste Domnul: Dupa cum pastorul scapa din gura leului numai doua bucati de picioare sau un vîrf de ureche, asa vor scapa copiii lui Israel care stau în Samaria în coltul unui pat si pe covoare de Damasc!\" (Amos 3:12).\r\n\r\nProfetia lui Amos nu este numai amenintare si pedeapsa. Dincolo de disciplinarea necesara, Amos vesteste recuperarea si restaurarea poporului. În numai cinci versete (Amos 9:10-15), viitorul Isarelului este descris în toata frumusetea împlinii legamintelor Avramic, Davidic si Palestinian în ceasul apoteotic al încununarii lui Mesia. În cartea lui Amos se gaseste textul profetic care, în vremea Bisericii primare, a pus capat dezbaterilor din consiliul din Ierusalim. Iacov este cel care-l tine minte si îl proclama tuturora: „ªi cu faptul acesta se potrivesc cuvintele proorocilor, dupa cum este scris: „Dupa aceea, Ma voi întoarce, si voi ridica din nou cortul lui David din prabusirea lui, îi voi zidi darîmaturile, si-l voi înalta din nou: pentru ca ramasita de oameni sa caute pe Domnul, ca si toate neamurile peste care este chemat Numele Meu, zice Domnul, care face toate aceste lucruri, si caruia Îi sînt cunoscute din vesnicie\" (Fapte 15:15-17 citat din Amos 9:11-12).\r\n\r\nBiblia ne îndeamna, iarasi si iarasi, sa nu uitam de Israel. Dumnezeu, chiar daca i-a pedepsit acum, nu i-a lepadat din planurile Sale. Evreii vor fi readusi în patria lor milenara. tara va înflori din nou. Natiunea va cunoaste iarasi prosperitatea si evreii pocaiti si reintrati în prerogativele Legamîntului vor redeveni martorii lui Dumnezeu în lume (Amos 9:13-15).\r\nI. OPT „GREUTATI\" 1 - 2\r\na. Damascul a invadat Israelul - 2 Împ. 10:32-33, (Amos 1:3)\r\nb. Gaza s-a aliat cu Tirul ca sa invadeze Iuda - 2 Cronici 21:16, 17;28:18, (Amos 1:6)\r\nc. Tirul s-a aliat cu Gaza ca sa invadeze Iuda, (Amos 1:9)\r\nd. Edom este dusmanos fata de Israel - Obadia 10-12, (Amos 1:11)\r\ne. Amon a atacat Galaadul, (Amos 1:13)\r\nf. Moab practica religii ucigatoare de oameni, (Amos 2:1)\r\ng. Iuda neglijeaza Legea Domnului - 2 Împ. 25:9, (Amos 2:4)\r\nh. Israel traieste În nelegiuire - 2 Împ. 17:17-23, (Amos 2:6)\r\n\r\nII. TREI PREDICI ÎMPOTRIVA ISRAELULUI 3 - 6\r\n1. Pedeapsa: justificata (Amos 3:1-10), pronuntata (Amos 3:11-15)\r\n2. Pedeapsa: justificata (Amos 4:1-11), pronuntata (Amos 4:12-13)\r\n3. Pedeapsa: justificata (Amos 5:1-15), pronuntata (Amos 5:16-6:14)\r\n\r\nIII. CINCI VEDENII 7 - 9\r\nLacustele, 7:1-3\r\nFocul, 7:4-6\r\nCumpana, 7:7-17\r\nCosul cu poame coapte, 8:1-14\r\nDumnezeu stînd pe altar, 9:1-10\r\nPromisiuni pentru viitorul lui Israel, 9:11-15" },
    //AMOS
    { "AMOS", "Explicatia Cartii: Amos\r\n\r\nIata o carte scrisa de un profet laic. Dumnezeu si-a pastrat întotdeauna dreptul sa se foloseasca de orice oameni pe care El îi alege pentru vestirea mesajelor Sale. Cînd preotii si împaratii îsi pierd vrednicia, ciobanii si culegatorii de smochine sînt înrolati în vestirea Cuvântului Domnului (Amos 7:14). Cartea poarta numele autorului ei: Amos, care se poate traduce prin „purtatorul de poveri\". Înca o data, numele este o anuntare a specificului misiunii pe care o va avea acest om chemat de Dumnezeu la o activitate grea si importanta.\r\nNumele lui Amos nu apare în nici o alta carte a Bibliei. El nu a fost un preot sau un profet educat în mînuirea si transmiterea Cuvîntului lui Dumnezeu. Ca si în cazul celorlalti profeti însa, autoritatea lui a venit de la Domnul si asta i-a fost de ajuns ca sa stea plin de îndrazneala în fata oamenilor:\r\n\r\n„Amos a raspuns lui Amatia: „Eu nu sînt nici prooroc, nici fiu de prooroc; ci sînt pastor, si strîngator de smochine de Egipt. Dar Domnul m-a luat de la oi, si Domnul mi-a zis: „Du-te si prooroceste poporului Meu Israel!\" (Amos 7:14-15).\r\n\r\nLocul din care s-a ridicat Amos a fost Tecoa, o asezare situata cam la 6 km sud de Betleem, în pustia iudeii. Acolo îsi pastorise altadata si David oile si tot acolo se ascunsese pe cînd fugea de împaratul Saul. Limbajul folosit de Amos este nesofisticat. Profetul are vorbirea unui om de la tara, plina de comparatii din natura în care-si duce viata (Amos 3:4, 5, 12; 5:8, 19; 9:9). Tonul lui este raspicat, iar mesajul lui merge direct la tinta fara prea multa pregatire: „El a zis: „Domnul racneste din Sion, glasul Lui rasuna din Ierusalim. Pasunele pastorilor jalesc, si vîrful Carmelului este uscat\" (Amos 1:2).\r\nAutorul fixeaza singur data activitatii sale: „pe vremea lui Ozia, împaratul lui Iuda si pe vremea lui Ieroboam, fiul lui Ioas, împaratul lui Israel, cu doi ani înaintea cutremurului de pamînt\". tinînd seama de amenintarile care ne apropie de sfîrsitul domniei lui Ieroboam, putem spune ca Amos a activat în preajma anului 755 î.Cr.\r\n\r\nAmos este un profet ridicat de Dumnezeu din împaratia lui Iuda si trimis sa prooroceasca la Betel si Samaria, în inima religioasa a împaratiei lui Israel. Activitatea profetica a lui Amos nu a fost solitara. Probabil ca în tineretea lui a mai apucat sa-l vada pe Elisei si sa-l auda pe Iona vorbind despre succesele militare ale Israelului (2 Împarati 14:25) ). Amos a fost premergatorul altor profeti. În timpul lucrarii lui, Dumnezeu l-a ridicat pe Osea sa vorbeasca în Israel si pe Isaia si Mica sa activeze în Iuda.\r\nAmos a profetit într-o vreme de optimism si abundenta. Vremea lui Ozia si Ieroboam readusese ceva din fala Israelului de altadata. La curtile împaratesti se instalase iarasi luxul, ritualurile religioase erau pline de pompa si fast, poporul traia iarasi în abundenta si în bucurie. Israelul arata însa ca un mar frumos pe dinafara si gaunos pe dinauntru. Pacatul se furisase în viata intima a poporului si dincolo de aparentele de religiozitate, Legea Domnului fusese abandonata. Simplitatea si neprihanirea fusesera înlocuite cu abundenta materiala, luxul, necinstea si imoralitatea (Amos 2:6-8; 3:10; 4:1; 5:10-12; 8:4-6). În plan international, Asiria, Mesopotamia si Egiptul erau în eclipsa, asa ca aparent, Israelul nu avea de ce sa se teama. În acest context Amos a fost trimis de Domnul sa predice un mesaj de mustrare si amenintare. Pedeapsa vestita de Amos nu se vedea atunci nici macar la orizontul istoriei, asa ca nu este de mirare ca acest „taran\" venit de la oi a fost privit cu curiozitate si cu aversiune în centrele civice si religioase ale Israelului. Imaginati-va cam cum au reactionat doamnele din „înalta societate\" a Samariei cînd l-au auzit spunîndu-le: „Ascultati acum, juncane din Basan, de pe muntele Samariei, voi care asupriti pe cei sarmani, zdrobiti pe cei lipsiti, si ziceti barbatilor vostri: „Dati-ne sa bem!\" (Amos 4:1). Nu este de mirare ca împotriva profetului s-a pornit repede o conspiratie care a facut totul ca sa-l reduca la tacere. Amatia, preotul idolatru din Betel, l-a parît împaratului, declarîndu-l tradator (Amos 7:10-11), iar apoi l-a interpelat public, gonindu-l din Israel:\r\n\r\n„Pleaca, vazatorule si fugi în tara lui Iuda! Manînca-ti pîinea acolo, si acolo prooroceste. Dar nu mai prooroci la Betel, caci este un locas sfînt al împaratului, si este un templu al împaratiei\" (Amos 7:12-13).\r\n\r\nAceasta interpelare a cauzat, declaratia autobiografica a profetului si groaznica sentinta dumnezeiasca rostita împotriva preotului Amatia (Amos 7:17). Alungat din Israel, Amos s-a întors în Iuda si a asternut continutul profetiilor sale în scris ca sa poata fi raspîndite în continuare în Israel si ca sa serveasca drept avertisment împaratiei lui Iuda.\r\nDesi de la tara, Amos este un vorbitor elocvent si ordonat. Capitolele 1 si 2 sînt o expunere despre „opt greutati\" care stau pe inima profetului. În revarsari profetice, Amos vorbeste împotriva tuturor celor opt natiuni din teritoriile Palestinei: Siria (1:3-5), Filistenii (1:6-8), Fenicia (1:9, 10), Edom (1:11-12), Amon (1:13-15). Moab (2:1-3), Iuda (2:4-5) si Israel (2:6-16). Urmeaza apoi capitolele 4, 5 si 6 care contin rezumatele a trei predici rostite de Amos. Începutul acestor trei mesaje este usor de gasit. Fiecare dintre ele încep cu: „Ascultati cuvîntul acesta...” (3:1; 4:1; 5:1). Ultimele capitole ale cartii cuprind o serie de 5 viziuni prin care, într-o succesiva accelerare a mîniei este vestita pedeapsa care va veni asupra celor ce s-au îndepartat de Domaul (Amos 7-9).\r\n\r\nCuvinte cheie si teme caracteristice: Doua expresii au pasit afara din cadrul cartii lui Amos înspre frazeologia tipologica crestina:\r\n\r\n„Pregateste-te sa te întîlnesti cu Dumnezeu tau\" (Amos 4:12) a devenit semnal de atentionare valabil oriunde si oricînd.\r\n\r\n„Vai de cei ce traiesc fara grija în Sion\" (Amos 6:1) a devenit sinonim cu avertizarea celor care se afla în pericolul formalismului religios.\r\nCartea lui Amos este bogata în mesaje spirituale. Ea raspunde întrebarii: „Cine poate sa slujeasca Domnului?\" cu un raspuns care trece dincolo de preotia traditionala, la masa mare a poporului, stabilind drept unic criteriu de calificare „chemarea Domnului pentru slujba\". Amos denunta ipocrizia si falsul ritualului religios lipsit de substanta reala a unei vieti traite în ascultare de Cuvîntul Domnului: „Duceti-va numai la Betel, si pacatuiti!... Aduceti-va jertfele în fiecare dimineata, si zeciuielile la fiecare trei zile! Faceti sa fumege jertfe de multumire facute cu aluat! Trimbitati-va, vestiti-va darurile de mîncare de buna voie! Caci asa va place, copii ai lui Israel, zice Domnul Dumnezeu\" (Amos 4:4-5).\r\n\r\nCartea lui Amos ne vorbeste despre iluzoria siguranta a belsugului si confortului zilei de astazi. Dumnezeu este acela care ridica sau coboara împaratii si împaratiile. Toate marile imperii ale lumii au cazut din cauza aceluiasi motiv groaznic: pacatul. El atrage ruina popoarelor. Peste Israelul prosper rasuna glasul de tunet al lui Amos: „Asa vorbeste Domnul: Dupa cum pastorul scapa din gura leului numai doua bucati de picioare sau un vîrf de ureche, asa vor scapa copiii lui Israel care stau în Samaria în coltul unui pat si pe covoare de Damasc!\" (Amos 3:12).\r\n\r\nProfetia lui Amos nu este numai amenintare si pedeapsa. Dincolo de disciplinarea necesara, Amos vesteste recuperarea si restaurarea poporului. În numai cinci versete (Amos 9:10-15), viitorul Isarelului este descris în toata frumusetea împlinii legamintelor Avramic, Davidic si Palestinian în ceasul apoteotic al încununarii lui Mesia. În cartea lui Amos se gaseste textul profetic care, în vremea Bisericii primare, a pus capat dezbaterilor din consiliul din Ierusalim. Iacov este cel care-l tine minte si îl proclama tuturora: „ªi cu faptul acesta se potrivesc cuvintele proorocilor, dupa cum este scris: „Dupa aceea, Ma voi întoarce, si voi ridica din nou cortul lui David din prabusirea lui, îi voi zidi darîmaturile, si-l voi înalta din nou: pentru ca ramasita de oameni sa caute pe Domnul, ca si toate neamurile peste care este chemat Numele Meu, zice Domnul, care face toate aceste lucruri, si caruia Îi sînt cunoscute din vesnicie\" (Fapte 15:15-17 citat din Amos 9:11-12).\r\n\r\nBiblia ne îndeamna, iarasi si iarasi, sa nu uitam de Israel. Dumnezeu, chiar daca i-a pedepsit acum, nu i-a lepadat din planurile Sale. Evreii vor fi readusi în patria lor milenara. tara va înflori din nou. Natiunea va cunoaste iarasi prosperitatea si evreii pocaiti si reintrati în prerogativele Legamîntului vor redeveni martorii lui Dumnezeu în lume (Amos 9:13-15).\r\nI. OPT „GREUTATI\" 1 - 2\r\na. Damascul a invadat Israelul - 2 Împ. 10:32-33, (Amos 1:3)\r\nb. Gaza s-a aliat cu Tirul ca sa invadeze Iuda - 2 Cronici 21:16, 17;28:18, (Amos 1:6)\r\nc. Tirul s-a aliat cu Gaza ca sa invadeze Iuda, (Amos 1:9)\r\nd. Edom este dusmanos fata de Israel - Obadia 10-12, (Amos 1:11)\r\ne. Amon a atacat Galaadul, (Amos 1:13)\r\nf. Moab practica religii ucigatoare de oameni, (Amos 2:1)\r\ng. Iuda neglijeaza Legea Domnului - 2 Împ. 25:9, (Amos 2:4)\r\nh. Israel traieste În nelegiuire - 2 Împ. 17:17-23, (Amos 2:6)\r\n\r\nII. TREI PREDICI ÎMPOTRIVA ISRAELULUI 3 - 6\r\n1. Pedeapsa: justificata (Amos 3:1-10), pronuntata (Amos 3:11-15)\r\n2. Pedeapsa: justificata (Amos 4:1-11), pronuntata (Amos 4:12-13)\r\n3. Pedeapsa: justificata (Amos 5:1-15), pronuntata (Amos 5:16-6:14)\r\n\r\nIII. CINCI VEDENII 7 - 9\r\nLacustele, 7:1-3\r\nFocul, 7:4-6\r\nCumpana, 7:7-17\r\nCosul cu poame coapte, 8:1-14\r\nDumnezeu stînd pe altar, 9:1-10\r\nPromisiuni pentru viitorul lui Israel, 9:11-15" },
    //OBADIA
    { "OBADIA", "Explicatia Cartii: Obadia\r\n\r\nCartea lui Obadia este stenograma unui proces în care este publicata rezolvarea unui conflict care a început între doi frati mai înainte ca ei sa se nasca. Cartea poarta un nume foarte comun în Israel: „Obadia\", care înseamna în traducere: „Slujitorul\" sau „Închinatorul lui Iehova\". Desi întîlnim în textul Bibliei cîtiva oameni care au purtat acest nume, este foarte improbabil ca vreunul dintre ei sa fie autorul acestei carti profetice.\r\nIdentitatea lui Obadia ramîne un mister. Textul nu ne da nici macar numele tatalui sau, ceea ce presupune ca autorul nu face parte în nici un fel din familiile preotesti sau regale din Israel. Unele date furnizate de textul cartii ne îndeamna sa credem ca Obadia a trait în împaratia lui Iuda.\r\nSingurul indiciu al vremii este una din invaziile împotriva evreilor. Acestea au fost multe si este dificil sa precizam la care dintre ele face aluzie textul. Aceasta precizare nici nu este importanta de fapt. Data scrierii nu este în nici un fel vitala pentru întelegerea si interpretarea cartii. Se prea poate ca aceasta carte sa fie cea mai veche dintre toate cartile profetice, iar Obadia sa fie unul dintre precursorii activitatii profetice din Israel.\r\nCartea lui Obadia este un rechizitoriu împotriva Edemului. Istoria acestui popor a început cu Esau, caruia i s-a mai spus si „Edom\" („Cel rosu\") din pricina „ciorbei rosiatice\" pentru care si-a vîndul lui Iacov dreptul de întîi nascut (Gen. 25:29-33). Edomitii erau urmasii acestui Esau care se stabilisera în muntele Sein „Esau s-a asezat în muntele Seir... Esau, tatal Edomitilor, în muntele Seir\" (Gen. 36:8, 9). Acest munte Seir nu era un pisc izolat, ci un întreg platou montan care se întindea de la sudul marii Moarte pîna la Golful Akaba. Numele regiunii venea de la conducatorul unui popor stravechi care populase acele locuri: „...si pe horiti în muntele lor, Seir..\", „Iata fii lui Seir, Horitul, vechi locuitori ai tarii...\" (Gen. 14:6; 36:20). Numele capitalei tinutului s-a chemat Sela sau Petra („stînca\"). Orasul era situat la extremitatea unui defileu foarte îngust, dar lung de aproape 2 kilometri si înalt de peste 200 de metri. Aceasta asezare geografica îi dadea un avantaj strategic extraordinar. O mîna de oameni putea apara intrarea în oras pe un termen foarte îndelungat. În depresiunea de la capatul defileului, locuitorii orasului Petra construisera peste l.000 de temple sapate în roca rosiatica. Locuintele oamenilor erau si ele sapate în stînca, parînd agatate în stînca precum cuiburile vulturilor (Obadia 4). În cartea Deuteronomul 2:12 gasim scris cum au devenit Edomitii locuitori ai muntelui Seir: „Seir era locuit altadata de Horiti; copiii lui Esau i-au izgonit, i-au nimicit dinaintea lor, si s-au asezat în locul lor\".\r\n\r\nÎntre Esau si Iacov, lupta a început înca din pîntecul mamei lor: „Copiii se bateau în pîntecele ei;... si Domnul i-a zis: „Doua neamuri sînt în pîntecele tau. ªi doua noroade se vor desparti la iesirea din pîntecele tau. Unul dintre noroadele acestea va fi mai tare decît celalalt. ªi cel mai mare, va sluji celui mai mic\" (Gen. 25:22-23). Înselat de Iacov, Esau, fratele mai mare, a plecat si s-a asezat în muntele Seir, dar dusmania dintre cei doi frati nu a încetat niciodata. Cînd urmasii lui Iacov au fost scosi de Dumnezeu din Egipt, ei au cerut voie urmasilor lui Esau sa-i lase sa treaca prin tinutul lor. Refuzul lor a fost însa categoric (Num. 20:14-21). Mai tîrziu, Edomitii i s-au împotrivit lui Saul. David si Solomon i-au subjugat pentru un timp. S-au rasculat iarasi împotriva evreilor în timpul lui Iosafat si Ioram (2 Cron. 21:8), au fost cuceriti iarasi de Amatia (2 Împ. 14:7) si s-au eliberat iarasi în timpul lui Ahab (2 Cron. 28:17). Dupa aceea, Edomitii au fost cuceriti de Nebucadnetar, împaratul Babilonului; în secolul cinci înainte de Cristos, tinutul lor a fost invadat de Nabateeni, si Edomitii au trebuit sa se refugieze în sudul Palestinei unde au fost cunoscuti ca \"idumeni\". Ironic, la nasterea Domnului Isus, un idumean, Irod, era pe tronul Israelului. Pentru a cîstiga încrederea si simpatia evreilor, el rezidise o parte a Templului. Încercarea lui de a-L ucide pe Cristos prin masacrarea pruncilor, este înca o dovada a acestui conflict dintre frati, întins peste veacuri. Idumenii au participat la rascoala Ierusalimului împotriva Romei imperiale (70 d.Cr.) si au fost nimiciti ca neam, limba si civilizatie imediat dupa aceea. Astfel s-a împlinit profetia rostita de Obadia: „Din pricina silniciei facute împotriva fratelui tau Iacov, vei fi acoperit de rusine si vei fi nimicit cu desavîrsire pentru totdeauna\" (Obadia 10).\r\nCartea lui Obadia este cea mai mica din întregul Vechi Testament. Textul ei este solemn si dur asemenea rostirii unei sentinte judecatoresti. La tribunalul cerului, Edom fusese gasit vinovat de împotrivire fata de planurile lui Dumnezeu si condamnat la distrugere vesnica. Vina lui Edom a fost ca n-a scapat nici un prilej de a face rau evreilor. Desi rude de sînge, dusmania dintre ei a determinat Edomul sa ia partea truturor celor care s-au repezit asupra lui Israel. Cartea este un verdict aspru împotriva unui popor care n-a stiut sa-si valorifice privilegiul de a fi iesit din aceiasi radacina cu Israelul. Condamnarea Edemului este contrastata în text cu viitorul luminos pe care-l promite Dumnezeu poporului Sau, pe care l-a ales ca sa mosteneasca „împaratia Domnului\" (Obadia 21).\r\nI. PEDEPSIREA EDOMULUI\r\na. Certitudinea ei, 1-9\r\nb. Motivele ei:\r\nLipsa iubirii de frate, 10\r\nIndiferenta, 11-12 A\r\ngresivitatea, 13-14\r\nc. Justificarea ei, 15-16\r\n\r\nII. PROMISIUNI PENTRU ISRAEL\r\na. Ascendent asupra Edomului, 17-19\r\nb. participare la împaratia Domnului, 18-21" },
    //IONA
    { "IONA", "Explicatia Cartii: Iona\r\n\r\nCartea profetului Iona a fost cea mai batjocorita carte a Bibliei. Elementele de interventie supranaturala din continutul ei au fost tinta ironiilor si discreditarilor din partea oamenilor de pseudo-stiinta. Cei care au cercetat însa cu atentie si cu credinta aceasta carte o aseaza printre cele mai frumoase carti care s-au scris cîndva. Cartea poarta numele autorului ei. „Iona\" se traduce prin „porumbel\", un nume potrivit pentru un crainic divin destinat sa fie purtator de vesti catre oameni.\r\nIona, fiul lui Amitai, s-a nascut în Gat-Hefer, o cetate aflata la 5 kilometri nord de Nazaretul Galileii. Singura referinta despre el pe care o mai gasim în Vechiul Testament este un text din 2 Împarati 14:25.\r\nTextul din 2 Împarati fixeaza activitatea proorocului Iona în timpul domniei lui Ieroboam al doilea (782-735 î.Cr.). Iona a fost un urmas al profetului Elisei si un precursor al profetilor Amos si Osea.\r\nTimpul în care a trait Iona a fost un timp de relativa prosperitate în Israel. Imperiul Asirian intrase pentru o vreme în declin, astfel ca Ieroboam al doilea a recucerit unele teritorii si a reîntors bunastarea în tara. Teama de cruzimea si puterea asiriana persista totusi în întreaga lume civilizata de atunci.\r\nCine citeste pentru prima data cartea lui Iona îsi pune foarte repede întrebarea: „Este aceasta carte înregistrarea unei întîmplari reale sau este o alegorie poetica?\" Pentru noi, autoritatea comentariului pe care-l face Domnul Isus despre cartea lui Iona este de ajuns sa ne convinga ca Iona a existat în realitate si ca patania lui a fost reala: „Un neam viclean si preacurvar cere un semn; dar nu i se va da alt semn, decît semnul proorocului Iona. Caci, dupa cum Iona a stat trei zile si trei nopti în pîntecele chitului, tot asa si Fiul omului va sta trei zile si trei nopti în inima pamîntului\" (Matei 12: 29-40). Cineva va putea spune ca Domnul a citat aceasta întîmplare tot asa cum citeaza cineva o ilustratie dintr-o carte de fictiune ca sa exemplifice un lucru real, dar aceasta neîncredere în realitatea istorica a întîmplarii este spulberata de un alt comentariu al Domnului: „Barbatii din Ninive se vor scula, în ziua judecatii, alaturi de neamul acesta, si-l vor osîndi, pentru ca ei s-au pocait la propovaduirea lui Iona; si iata ca aici este Unul mai mare decît Iona\" (Luca 11:32; Matei 12:41).\r\n\r\nMulti s-au poticnit pentru ca în cartea lui Iona scrie ca profetul a fost înghitit de un peste mare si ca a iesit de acolo dupa trei zile viu. Arhivele navale dau însa dovezi clare ca astfel de întîmplari sînt posibile si ca s-au repetat chiar de mai multe ori în decursul vremii. Ziarul „Daily Mail\", aparut la Londra în 14 Decembrie 1928, descrie dimensiunile unui soi de peste marin adus împaiat la o expozitie. Oamenii erau îndemnati sa urce o scara si sa se catere prin gura pestelui în stomacul lui spatios.\r\n\r\nSir Francisc Fox scrie în cartea sa: „63 de ani de inginerie\" despre cazul unui casalot în stomacul caruia s-a gasit scheletul unui rechin lung de nu mai putin de 5 metri.\r\n\r\nDomnul M. de Parville, editorul stiintific al ziarului francez „Journal de debats\", povesteste cum, în timpul vînarii unui casalot, doi marinari au cazut în apa oceanului. Unul a fost salvat, în timp ce celalalt disparuse fara urma. Cînd au reusit sa traga în sfîrsit uriasul casalot pe punte si i-au despicat stomacul, spintecatorii l-au gasit înauntru pe Bartley, marinarul disparut. Pielea lui fusese arsa de acidul sucului gastric, dar el si-a revenit repede la viata. Medicul a constatat ca Bartley lesinase de frica, nu din pricina lipsei de aer.\r\n\r\nÎn luna Noiembrie, ziarul „Mail\" din Madras, India, relata despre prinderea unui rechin de dimensiuni uriase în pîntecele caruia a fost descoperit scheletul unui om si hainele care-i atîrnau înca pe oase. S- a facut precizarea ca rechinul fusese prins la 50 de kilometri de Bombei, iar victima fusese probabil unul dintre cei disparuti recent într-un taifun.\r\n\r\nExemplele ar putea continua, dar ce rost au ele în contextul unor afirmatii atotsuficiente ale Domnului Isus Cristos, stîlpul si temelia adevarului?\r\n\r\nRevenind la cartea lui Iona, este bine sa spunem ca ea ne vorbeste despre doua teme paralele: (1) un profet aflat în scoala lui Dumnezeu si (2) un Dumnezeu caruia îi place sa se îndure de oameni.\r\n\r\nÎntîmplarea este simpla, dar graitoare: Iona este trimis cu un mesaj de pocainta la cetatea pagîna Ninive. Profetul cunostea însa rautatea asirienilor si pericolul potential pe care-l constituiau ei pentru evrei (Isaia profetise cu aproximativ 30 de ani mai înainte despre nenorocirile pe care asirienii le vor aduce asupra celor din Israel - Isaia 7:17-25) si din aceasta cauza el refuza sa plece la Ninive si fuge „departe de fata Domnului\" cu o corabie care se îndrepta spre Tars. Dumnezeu îl opreste din fuga lui razvratita prin intermediul unei furtuni stîrnita pe mare, trimite un peste sa-l scape pe Iona de la înec si sa-l întoarca pe tarmul marii si-l trimite pe profet a doua oara la Ninive. Rezultatul propovaduirii lui Iona este pocainta ninivenilor, dar si mînia profetului, care-i vede pe dusmanii poporului sau scapati de pedeapsa divina. Urmeaza o lectie plina de tact pe care Dumnezeu i-o da lui Iona, ca sa-l faca sa cunoasca ceva din dragostea pe care Creatorul o simte pentru toate creaturile Sale. Nationalismul profetului este înnecat în imensitatea iubirii divine. Iona accepta noua revelatie a lui Dumnezeu si se asterne la scris pentru a ne lasa noua aceasta extraordinara lectie de disciplinare a unui profet, care si-a iubit prea mult neamul si prea putin pe Dumnezeul sau si pe oamenii din popoarele înconjuratoare.\r\n\r\nCele 4 capitole ale cartii sînt tot atîtea scene care ar putea purta niste titluri de sine statatoare: (1) Iona si furtuna, (2) Iona si pestele, (3) Iona si orasul Ninive si (4) Iona si Domnul. Rînd pe rînd îl vedem pe Iona fugind departe de fata Domnului, rugîndu-se Domnului, predicînd în numele Domnului si învatînd de la Domnul. Cele 4 capitole ne arata succesiv: neascultarea, recuperarea, folosirea si educarea lui Iona.\r\n\r\nCuvinte cheie sl teme caracteristice: Exista 5 lucruri pe care le foloseste Dumnezeu pentru educarea spirituala a profetului:\r\n\r\n(1). furtuna stîrnita pe mare, (Iona l:4)\r\n(2). un peste mare, (Iona 1:17)\r\n(3). umbrarul de curcubete, (Iona 4:6)\r\n(4). un vierme mic, (Iona 4:7)\r\n(5). un vînt uscat si firbinte din rasarit, (Iona 4:8)\r\n\r\nPatania lui Iona intra în categoria „tipurilor\" profetice. Iona simbolizeaza trei lucruri:\r\n\r\na. Destinul si atitudinea Israelului în planul lui Dumnezeu de mîntuire a întregii lumi. În peripetiile lui Iona vedem simbolic peregrinarile evreilor. ªi ei au fost neascultatori de Dumnezeu în împlinirea chemarii lor de a fi lumina neamurilor. ªi ei au cautat scapare printre neamuri. ªi ei au adus nenorocirea asupra neamurilor, asa cum Iona a atras furtuna asupra marinarilor. ªi ei L-au vestit pe Dumnezeul lor în mijlocul crizelor produse între neamuri, tot asa cum Iona L-a vestit pe Dumnezeu marinarilor. ªi ei au fost pedepsiti de neamuri fiind considerati mereu un fel de pricina a tuturor rautatilor si calamitatilor. ªi ei au fost pastrati în mod miraculos de Dumnezeu. ªi ei îsi vor relua curînd misiunea lor, care va produce pocainta si mîntuirea neamurilor (Zaharia 8:13, 20 si mai ales 23: „În zilele acelea, zece oameni din toate neamurile vor apuca pe un iudeu de poala hainei, si-i vor zice: „Vrem sa mergem cu voi; caci am auzit ca Dumnezeu este cu voi.\")\r\n\r\nb. Înmormîntarea si învierea Domnului Isus. Lucrul acesta s-a vazut clar în declaratiile Domnului Isus înregistrate de Evanghelii. De ce a trebuit ca Iona sa stea neaparat trei zile si trei nopti în pîntecele pestelui? Pentru ca atît era planuit sa stea Domnul Isus în mormînt înainte de înviere.\r\n\r\nc. Viata si propovaduirea Domnului Isus însusi. Textul din Evanghelia lui Matei ne spune ca Iona a fost un semn pentru Niniveni. Fara îndoiala ca el le-a spus ceva despre fuga lui de Domnul si de întîmplarea cu pestele. Se prea poate ca pielea de pe trupul lui sa fi vorbit de la sine dovedind veridicitatea spuselor profetului. Impresia pe care a produs-o Iona asupra poporului din Ninive a fost coplesitoare. Tot asa si Domnul Isus este astazi un semn pentru mîntuirea neamurilor.\r\n\r\nMesajul peste veacuri: Capitolul 4 al cartii este si momentul ei culminant. Pentru cititorul neavizat, cartea lui Iona se întrerupe abrupt si nenatural. Pentru cel familiarizat cu tehnica literaturii ebraice însa stie ca versetele din final sînt de fapt tema întregii carti si ca de fapt pentru enuntarea lor a fost cladit întregul esafodaj al cartii:\r\n\r\n„tie îti este mila de curcubetele acesta, care nu te-a costat nici o truda si pe care nu tu l-ai facut sa creasca, ci într-o noapte s-a nascut si într-o noapte a pierit. ªi Mie sa nu-Mi fie mila de Ninive, cetatea cea mare, în care se afla mai mult de o suta douazeci de mii de oameni, care nu stiu sa deosebeasca dreapta de stînga lor, afara de o multime de vite!\" (Iona 4:10-11).\r\n\r\nAceasta descoperire a launtrului inimii lui Dumnezeu se ridica mai sus decît majoritatea documentelor sfinte din arhivele evreilor, atingînd culmea Nou Testamentala a afirmatiei din Ioan 3:16:\r\n\r\n„Fiindca atît de mult a iubit Dumnezeu lumea, ca a dat pe singurul Sau Fiu, pentru ca oricine crede în El, sa nu piara, ci sa aiba viata vesnica\".\r\n\r\nÎn misiunea lui Iona recunoastem prefatarea misiunii pe care o vor avea peste veacuri ucenicii Domnului Isus:\r\n\r\n„Mergeti în toata lumea si propovaduiti Evanghelia la orice faptura. Cine va crede si se va boteza, va fi mîntuit; dar cine nu va crede va fi osîndit\" (Marcu 16:15-16).\r\nI. FUGA LUI IONA\r\na. Motivele fugii lui, 1:1-2\r\nb. Itinerarul fugii lui, 1:3\r\nc. Rezultatul fugii lui, 1:4-17\r\n\r\nII. RUGaCIUNEA LUI IONA\r\na. Continutul rugaciunii lui, 2:1-9\r\nb. Consecinta rugaciunii lui, 2:10\r\n\r\nIII. PREDICA LUI IONA\r\na. Trimiterea divina, 3:1-3\r\nb. Mesajul predicat de Iona, 3:4\r\nc. Efectul predicarii lui Iona, 3:5-10\r\n\r\nIV. LECtIA PRIMITa DE IONA\r\na. Nemultumirea lui Iona, 4:1-3\r\nb. Explicatia pe care i-o da Domnul, 4:4-11" },
    //MICA
    { "MICA", "Explicatia Cartii: Mica\r\n\r\n„ti s-a aratat, omule, ce este bine, si ce alta cere Domnul de la tine, decît sa faci dreptate, sa iubesti mila, si sa umbli smerit cu Dumnezeul tau?\" (Mica 6:8) Cartea poarta numele autorului ei: Mica. Numele acesta este o forma prescurtata a lui „Mkaiahu\" care înseamna în traducere: „Cine este ca Iehova?\". Profetul face un joc de cuvinte, ca o veritabila semnatura în finalul profetiei: „Cine este ca Tine...” (Mica 7:18).\r\nMica a trait în Moreset, la aproximativ 35 de kilometri sud-vest de Ierusalim (Mica 1:4). El si-a început activitatea profetica la putina vreme dupa ce Isaia si-o începuse pe a sa. Cei doi profeti din împaratia lui Iuda au fost fara îndoiala prieteni si si-au împartasit unul altuia multe din framîntarile si preocuparile lor. Isaia, fiind de neam ales, a profetit la curtea împarateasca. Mica a fost un om din popor si si-a desfasurat activitatea printre oamenii simpli. Dumnezeu a vorbit însa la fel si prin unul si prin celalalt. Exista o izbitoare asemanare între anumite pasaje din cartile celor doi. (Mica 4: l -5 si Isaia 2:2-4).\r\nPrimul verset al cartii ne informeaza ca Mica a profetit pe vremea lui Iotam (739-731 î.Cr.), Ahaz (731-715 î.Cr.) si Ezechia (715-686 î.Cr.). Activitatea profetica a lui Mica trebuie deci sa se fi desfasurat între anii 735 si 710 î.Cr. În afara de Isaia, el a mai fost contemporan si cu Osea, care profetea în împaratia lui Israel.\r\nCu toate ca, în general, nu a fost un împarat rau, Iotam nu a îndepartat înaltimile idolatre din împaratia lui Iuda. Ahaz, care i-a urmat a fost un împarat rau (vezi 2 Împ. 16:2-4). El a tradat încrederea în Dumnezeu si a adoptat o politica pro-asiriana, în timp ce el domnea în Iuda, cele 10 semintii din împaratia de nord au fost duse în robie. Dupa Ahaz, a urmat Ezechia, care a fost un împarat surprinzator de bun. El i-a înfruntat pe asirieni si a rezistat în mod miraculos unui puternic asediu (2 Împarati 18:13-19:36).\r\n\r\nPentru oamenii de la tara, timpurile acelea au fost timpuri grele. Navalitorii amenintau mereu cu invadarea pamînturilor, bogatii tarii exploatau fara mila poporul (Mica 2:1-l 3), iar conducatorii lui Iuda depasisera orice limita si „îi mîncau si-i jupuiau pe saracii tarii\" (Mica 3: 1-4). Ca si Amos, Mica s-a ridicat plin de curaj si a strigat împotriva nedreptatii.\r\nO treime din cartea lui Mica denunta pacatele concetatenilor lui; o alta treime descrie pedeapsa pe care o va trimite Dumnezeu asupra lor; iar ultima parte dezvaluie nadejdea profetului în restaurarea pe care o va lucra Dumnezeu în tara si slava de care va fi urmata. Desi este din Iuda, Mica profeteste si despre pedepsirea Samariei (Mica l:6).\r\n\r\nCuvinte cheie si teme caracteristice: Capitolele 6 si 7 descriu judecata lui Dumnezeu cu poporul evreu. Curtea de judecata este formata din muntii si dealurile pamîntului. Teatralitatea pasajului reuseste sa scoata si mai bine în evidenta vinovatia poporului înaintea dreptatii divine.\r\n\r\nDesi anunta inevitabilitatea pedepsei, cartea scoate în evidenta clementa pe care o manifesta Dumnezeu, ca Judecator suprem, fata de Israel: „Care Dumnezeu este ca Tine, care ierti nelegiuirea, si treci cu vederea pacatele ramasitei mostenirii Tale? El nu tine mînia pe veci, ci îi place îndurarea!\"(Mica 7:18). Dumnezeu a stiut întotdeauna sa crute o ramasita a lui Israel pentru a-ªi putea continua împlinirea planurilor Sale mesianice.\r\n\r\nAsemenea tuturor celorlalti profeti, Mica subliniaza harul de care se bucura evreii în virtutea legamintelor încheiate de Dumnezeu cu ei: „El va avea iarasi mila de noi, va calca în picioare nelegiuirile noastre, si vei arunca în fundul marii toate pacatele lor. Vei da cu credinciosia lui Iacov, si vei tinea cu îndurare fata de Avraam, ce ai jurat parintilor nostri în zilele de odinioara\" (Mica 8:19-20).\r\nMica priveste peste orizontul timpului, anunta întruparea Fiului lui Dumnezeu în mica cetate a Betleemului si proclama domnia pe care El o va exercita asupra pamîntului: „ªi tu, Betleeme Efrata, macar ca esti prea mic între cetatile de capetenie a lui Iuda, totusi din tine îmi va iesi Cel ceva stapîni peste Israel, si a carui obîrsie se suie pîna în vremuri stravechi, pîna în zilele vesniciei. De aceea îi va lasa pîna va naste cea care are sa nasca, si ramasita fratilor Sai se va întoarce la copiii lui Israel. El se va înfatisa si va cîrmui cu puterea Domnului, Dumnezeului Sau: vor locui linistiti, caci El va fi proslavit pîna la marginile pamîntului. El va fi pacea noastra!\" (Mica 5:2-5).\r\nI. ANUNTAREA PEDEPSEI 1 - 3\r\na. Descrierea pedepsei, 1:2-16\r\nb. Motivarea pedepsei, 2:1-13\r\nc. Destinatarii pedepsei:\r\nCapeteniile lui Iacov, 3:1-4\r\nProfetii mincinosi, 3:5-8\r\nCetatea Ierusalim, 3:9-12\r\n\r\nII. ANUNtAREA SLAVEI VIITOARE 4 - 5\r\na. Slava împaratiei, 4:1-8\r\nb. Suferintele curatitoare, 4:9-5:1\r\nc. Aratarea lui Mesia\r\nLa prima lui venire, 5:2-3\r\nLa cea de a doua venire a Lui, 5:4-15\r\n\r\nIII. ÎNDEMN LA POCaINta 6 - 7\r\na. Primul cap de acuzare, 6:7-5\r\nb. Primul raspuns de aparare, 6:6-7\r\nc. Al doilea cap de acuzare, 6:8-16\r\nd. Al doilea raspuns de aparare, 7:1-10" },
    //NAUM
    { "NAUM", "Explicatia Cartii: Naum\r\n\r\nProfetia lui Naum este asemanatoare unui poem cinematografic. Cartea nu transmite idei prin discursuri, ci printr-o succesiune ametitoare de imagini. Numele lui Naum înseamna în limba ebraica: „mîngîiere\". Numele orasului Capernaum, în care s-a stabilit pentru o vreme Domnul Isus se traduce prin: „Cetatea mîngîierii\". Mesajul lui Naum a adus mîngîiere tuturor celor ce traiau sub groaznica dominatie asiriana. Imperiul asirian a fost unul dintre cele mai crude din istorie. Jupuirea oamenilor de vii, smulgerea limbii sau taierea mîinilor si picioarelor erau numai cîteva din pedepsele celor ce îndrazneau sa nu se supuna Asiriei sau care încercau sa se rascoale. În vremea lui Naum, Asiria era sinonima cu teroarea. Vestea despre apropiata ei pedepsire venea ca o briza racoritoare peste inimile si mintile arse de teama si suferinta.\r\nSingura informatie despre Naum ne-o pune la dispozitie textul din prefata profetiei lui: „Naum, din Elcos\" (Naum l:1). Nu se stie cu siguranta unde s-a aflat aceasta localitate.\r\nTextul din Naum 3:8-10 vorbeste despre caderea Tebei (No-Amon) ca despre un eveniment petrecut recent, deci cartea trebuie sa fie scrisa cu putin timp dupa anul 664 î.Cr. Cetatea Teba a fost rezidita 10 ani mai tîrziu, asa ca data scrierii nu poate fi plasata dupa anul 654 î.Cr.\r\nPocainta ninivenilor sub efectul propovaduitii profetului Iona nu a fost de prea lunga durata. Asiria s-a întors curînd la practicile ei vinovate. În anul 722 î.Cr., teama lui Iona s-a dovedit a fi fost întemeiata, caci Sargon al doilea a venit sa cucereasca si sa distruga Samaria, ducînd în robie cele 10 semintii ale lui Israel. Pe vremea lui Naum (circa 660 î.Cr.), Asiria s-a aflat la apogeul puterii si dominarii ei imperiale. Sa prevezi în acel moment sfîrsitul Asiriei, era un act de curaj si de intuitie pe care numai un om cu acces la descoperirile dumnezeiesti asupra viitorului o putea face. Naum a fost un astfel de om cu capacitati supranaturale. Profetii din vechime mai erau numiti si „vazatori\". Nimeni n-a purtat vreodata mai pe merit aceasta numire ca Naum. Vedeniile „cinematografice\" primite de el de la Dumnezeu au functionat ca un veritabil „tunel al timpului\" prin care profetul a fost proiectat ca martor ocular al unor evenimente istorice care urmau sa se desfasoare peste alti multi ani de zile.\r\n\r\nSub Asurbanipal (669-633 î.Cr.) Asiria si-a atins gloria existentei ei. Capitala imperiului, Ninive, devenise unul dintre cele mai fortificate si mai marete orase ale lumii. De jur împrejur, zidurile erau înalte de 35 de metri si îndeajuns de late pentru ca trei care de lupta sa poata trece în paralel. Din loc în loc, zidul era prelungit în înaltime cu turnuri de înca 30 de metri, în afara fortificatiilor zidului, cetatea mai era înconjurata si de un sant de apa lat de 50 de metri si adînc de 20 de metri. Ninive parea a fi imposibil de cucerit. Între zidurile ei, populatia acumulase provizii pentru un asediu de 20 de ani de zile. La vremea aparitiei ei, proorocia rostita de Naum parea un vis frumos, dar irealizabil. Totusi, Dumnezeu a facut întocmai asa cum anuntase prin vocea profetului. Naum anuntase ca Ninive va pieri din cauza unor „valuri ce se revarsa peste mal\" (Naum l:8). ªi asa a si fost. În timpul asediului asezat de babilonieni, Eufratul s-a umflat din cauza ploilor si a rupt o parte din zidul pe sub care intra în cetate. Ostile cuceritorilor au intrat prin gaura aceea, i-au nimicit pe locuitorii cetatii, au jefuit tot ceea ce se putea jefui si apoi au pus foc (Naum 3:13, 15). Naum vestise de asemenea ca Ninive va fi „ascunsa\" (Naum 3:11). Ruinele ei, uitate de timp, au fost redescoperite abia în anul 1842 d.Cr., facîndu-i înca o data sa taca pe cei ce timp de secole batjocorisera competenta istorica a Bibliei.\r\nProfetia lui Naum este un reportaj la fata locului despre distrugerea cetatii Ninive. Condamnarea si pedepsirea ei pot fi o pilda pentru judecata pe care a rostit-o Dumnezeu împotriva întregii lumi pacatoase si razvratite. Proclamatia lui Naum despre soarta cetatii Ninive, nu mai este o chemare la pocainta, ca pe vremea lui Iona, ci publicarea unui verdict ireversibil.\r\n\r\nCartea lui Naum este un excelent studiu în caracterul lui Dumnezeu. Atributele Lui sînt prezentate în toata splendoarea si echilibrul lor. În textul din Naum l:2, 3, 6, 7, gasim pe de o parte trasaturile mîniei lui: „Domnul este un Dumnezeu gelos si razbunator... plin de mînie... tine mînie pe vrajmasii Lui... cine poate sta împotriva urgiei Lui?\" iar pe de alta trasaturile bunatatii Lui vesnice: „Domnul este îndelung rabdator... Domnul este bun, El este un loc de scapare în ziua necazului; si cunoaste pe cei ce se încred în El\".\r\n\r\nIona ne-a aratat un Dumnezeu al iubirii si al harului, Naum ne arata un Dumnezeu al sfinteniei, „de o mare tarie, si (care) nu lasa nepedepsit pe cel rau\" (Naum l:3).\r\nMesajul peste veacuri: „Ce seamana omul, aceea va culege\" spune proverbul Biblic. Cartea lui Naum este un avertisment pentru toti aceia care se împietresc în timpul îndelungii rabdari a lui Dumnezeu, crezînd ca pedepsirea lor nu va veni niciodata: „Unde este fagaduinta venirii Lui? Caci de cînd au adormit parintii nostri, toate ramîn asa cum erau de la începutul zidirii!\" (2 Petru 3:4). Ninive este un „tip\" pentru toate acele neamuri care îi întorc spatele lui Dumnezeu. Puternicii lumii si-au închipuit mereu ca pentru o îndelunga dominatie este nevoie de putere militara si economica. Profetia lui Naum ne arata însa ca Ninive s-a prabusit din cauza pacatului (Naum 3:1-7) si ca în caderea ei, cetatea nu s-a putut sprijini nici pe bogatia economica si nici pe taria militara (Naum 3:8-19). Persoana sau natiunea care în mod deliberat si hotarît îl refuza pe Domnul, se îndreapta deliberat si hotarît spre distrugere. Mesajul acesta a fost necesar atunci si ramîne necesar si astazi!\r\nI. MARETIA LUI DUMNEZEU\r\na. Maretia atributelor Sale, 1:2-8\r\nb. Marimea mîniei Sale, 1:9-14\r\n\r\nII. JUDECATA DIVINA\r\na. Provocarea, 2:1-2\r\nb. Pustiirea, 2:1-3:19" },
    //HABACUC
    { "HABACUC", "Explicatia Cartii: Habacuc\r\n\r\nCartea lui Habacuc este o cîntare în noapte. Inima omului este un nesfîrsit subiect de studiu. Nu exista doi oameni la fel, si deci nici doi prooroci identici. Habacuc este unic în felul în care reactioneaza în fata revelatiei divine. Aceasta carte nu este un mesaj catre oameni, ci un dialog între mintea profetului care dorea sa înteleaga si inima lui Dumnezeu care nu vrea întotdeauna sa explice. Titlul cartii este numele autorului ei: „Habacuc\" si se poate traduce prin „cel ce tine în brate\".\r\nHabacuc a fost contemporan cu Ieremia si a activat în perioada de dinaintea ducerii împaratiei lui Iuda în robia Babiloniana. Profetul a fost un om cu un caracter sensibil si duios care a îndeplinit slujba de „strajer\" spiritual pus de Dumnezeu peste poporului Sau („M-am dus în locul meu de straja, si stam în turn ca sa veghez si sa vad ce are sa-mi spuna Domnul\" - Hab. 2:1).\r\nHabacuc a activat ca profet în secolul 7 dinaintea lui Cristos, în perioada premergatoare robiei babiloniene.\r\nCu ani de zile înainte de aparitia lui Habacuc, profetul Isaia îl anuntase pe Ezechia, împaratul lui Iuda ca toata visteria împaratiei lui va fi luata si dusa în Babilon (Isaia 39:6,7). Pe vremea aceea însa, toata lumea se temea de Asiria. Babilonul a început sa constituie o amenintare numai dupa caderea cetatii Ninive. Cînd Egiptul a venit sa dea o mîna de ajutor la asedierea si cucerirea cetatii Ninive, Iosia, care era un supus al Asiriei, i-a iesit în cale sa-l opreasca din drum si a fost ucis în lupta de la Meghido (2 Împarati 23:28-30). Fiul lui Iosia, Ioahaz, a apucat sa domneasca numai trei luni, caci a fost luat „ostatec\" în Egipt de faraonul Neco, care a asezat pe tron în locul lui pe celalalt fiu al lui Iosia, Eliachim (2 Împarati 23:31-37). Neco i-a schimbat numele în Ioiachim. Acesta a fost împaratul de trista amintire care a ars sulul cartii lui Ieremia (Ieremia 36). Pe vremea lui, babilonienii s-au înfruntat cu egiptenii si i-au învins la Carchemis, lînga rîul Eufrat. Urmarindu-i pe egiptenii care se retrageau, Nebucadnetar a asediat si Ierusalimul si l-ar fi cucerit fara nici o îndoiala, daca vestea mortii tatalui sau, Nebopolazar, nu l-ar fi fortat sa se întoarca de urgenta acasa. Asa cum anuntase Habacuc, Babilonul devenise o amenintare pentru Iuda.\r\nCartea lui Habacuc este un unicat între cartile profetice. Ea nu este un mesaj catre popor, ci o discutie dintre profet si Dumnezeul care-l ridicase în slujba. Textul ei nu anunta prea multe evenimente, ci cauta sa rezolve o problema. Habacuc era încurcat cu dreptatea hotarîrilor lui Dumnezeu în istorie. Mintea lui nu putea întrezari justetea hotarîrilor divine. Confuz si curios, el cere o audienta si un raspuns de la Domnul.\r\n\r\nCartea ne prezinta un om care, desi se încrede în Dumnezeu, este perplex în fata hotarîrilor Lui. Dilema lui Habacuc era de o dubla natura: (1) De ce îngaduie Dumnezeu ca nelegiuirea crescînda din Iuda sa ramîna nepedepsita? (Hab. 1:2-4) si (2) Cum poate un Dumnezeu sfînt si drept sa se gîndeasca sa pedepseasca vinovata împaratie a lui Iuda prin si mai vinovatul Babilon? (Hab. 1:12-2:1). Raspunsurile pe care i le da Dumnezeu la aceste doua întrebari sînt consemnate în Hab. 1:5-11 si Hab. 2:2-20. Cartea se încheie apoi cu o oda de lauda la adresa suveranitatii si atotîntelepciunii divine. Textul cîntarii de lauda este unul dintre cele mai frumoase din Biblie si a fost asezat de evrei pe muzica si cîntat la serviciile de la Templu.\r\n\r\nCuvinte cheie si teme caracteristice: Are voie un om sa chestioneze justetea hotarîrilor lui Dumnezeu? Aceasta este problema pe care o ridica profetia lui Habacuc. Raspunsul pe care-l da cartea este ca omul are dreptul sa întrebe, dar raspunsul pe care-l va primi de la Domnul va mari uneori si mai mult neîntelegerea omului. Mintea noastra limitata în circumstantele de timp si spatiu nu poate patrunde justetea planurilor lui Dumnezeu în istorie. Suferintele neîntelese de astazi vor fi cu siguranta explicate de Dumnezeu „în ziua aceea\". Pîna atunci, ne spune Habacuc, „cel neprihanit va trai prin credinta\" (Hab. 2:4). Tema cuprinsa în principiul acesta este temelia pe care a fost construita mîntuirea oferita în Noul Testament. Habacuc 2:4 este citat de 3 ori în texte doctrinare fundamentale din epistole (Rom 1:17; Gal. 3:11; Evrei 10:38).\r\n\r\nUn alt lucru important din cartea lui Habacuc este atitudinea proorocului atunci cînd are o problema. El nu merge la oameni pentru rezolvare, ci se retrage în singuratatea partasiei cu Dumnezeu asteptînd „în tacere ajutorul Domnului\". O astfel de atitudine este întotdeauna rasplatita de Domnul. Nu este adevarat ca Dumnezeu nu mai vorbeste ca altadata! Este însa adevarat ca oamenii nu mai stiu sa asculte ca în alte timpuri! „Domnul însa este în Templul Lui cel sfînt. Tot pamîntul sa taca înaintea lui\" (Hab. 2:20).\r\nI. UN PROFET CU PROBLEME\r\na. Problema #1:\r\nDe ce îngaduie Dumnezeu propasirea raului? 1:2-4\r\nb. Raspunsul, 1:5-11\r\nc. Problema #2:\r\nCum poate folosi Dumnezeu un neam mai rau pentru pedepsirea raului din Iuda?, 1:12-2:1\r\nd. Raspunsul, 2:2-20\r\n\r\nII. ODA CREDINtEI\r\na. Lauda pentru persoana Domnului, 3:1-2\r\nb. Lauda pentru puterea Domnului, 3:4-7\r\nc. Lauda pentru scopurile Domnului, 3:8-16\r\nd. Lauda asezata pe temelia credintei, 3:17-19" },
    //ȚEFANIA
    { "ȚEFANIA", "Explicatia Cartii: Tefania\r\n\r\n„Voi nimici totul de pe fata pamîntului, zice Domnul. Tacere înaintea Domnului Dumnezeu! Caci ziua Domnului este aproape, caci Domnul a pregatit jertfa, ªi-a sfintit oaspetii\" (tef. 1:2,7) Numele profetului, care formeaza si titlul cartii vine de la doua cuvinte ebraice: „tefan-Iah\" care se pot traduce prin: „Dumnezeu care ascunde\".\r\ntefania este singurul profet care merge cu propria linie genealogica pîna la a patra generatie: „tefania, fiul lui Cusi, fiul lui Ghedalia, fiul lui Amaria, fiul lui Ezechia” (tef. 1:1). Motivul acestei prezentari este acela al sublinierii apartenentei lui la linia davidica. tefania a fost un print din familia lui David care a fost chemat de Dumnezeu la slujba profetiei. El a fost deci ruda de sînge cu împaratul Iosia, în timpul caruia si-a desfasurat activitatea. Se prea poate ca influenta lui tefania sa fi fost unul dintre factorii care l-au influentat pe Iosia în miscarea lui de trezire spirituala a tarii.\r\ntefania ne da el însusi datele lucrarii lui profetice: „pe vremea lui Iosia, fiul Amon, împaratul lui Iuda\" (tef. 1:1). tinînd seama ca acest Iosia a domnit între 640-609 î.Cr. si ca Ninive nu fusese înca nimicita (tef. 2:13), putem fixa scrierea acestor profetii între anii 640-612 î.Cr., aceasta însemnînd ca tefania l-a însotit pentru o vreme pe Ieremia în începutul misiunii sale.\r\ntefania este profetul nemultumit în vreme de trezire spirituala. Domnia lui Iosia a adus în Iuda o vreme de întoarcere la litera si ritualul religios din „vremurile bune\". Prima reforma a lui Iosia s-a produs în anul 628 î.Cr. în cel de al doisprezecelea an de domnie (2 Cron. 34:3-7). Împaratul a înlaturat altarele lui Baal, a darîmat stîlpii idolesti si a ars oasele profetilor mincinosi pe altarele lor. ªase ani mai tîrziu (622 î.Cr.), cea de a doua miscare de reforma condusa de Iosia s-a produs atunci cînd preotul Hilchia a gasit Cartea Legii în Templu (1 Cron. 34:8-35:19). Aceste doua reforme au fost însa prea mici si au venit prea tîrziu. Schimbarea a fost mai mult administrativa, decît spirituala (2 Imp. 22-23) Proorocul Hulda i-a spus lui Iosia ca Dumnezeu apreciaza intentiile lui bune, dar ca poporul neînfrînat si neschimbat în trairile lui va trebui pedepsit cu asprime (2 Împ. 22:15-20). Înnoirea facuta în tara de Iosia a fost mai mult o reforma, decît o renastere. Acest lucru a fost proclamat cu tarie de Ieremia: vicleana Iuda, nu s-a întors la Mine din toata inima ei, ci cu prefacatorie\" (Ieremia 3:6-10).\r\nNu este de mirare deci ca tefania nu pomeneste nimic despre reformele religioase ale lui Iosia. Analiza lui asupra starii spirituale a poporului a mers dincolo de straturile superficiale ale aparentelor. Ceea ce a vazut el în starea evreilor a fost nestatornicie, pacat si nelegiuire. Expunînd public vinovatia oamenilor, tefania striga cu putere anuntînd venirea „zilei Domnului\". Acest uragan al mîniei divine se va napusti peste tara si o va face una cu pamîntul. tefania si Ioel au cîntat împreuna oda razbunarii lui Dumnezeu asupra împaratiei lui Iuda.\r\n\r\nDincolo de vremea pedepsei, tefania vede si el vremea îndurarii si a slavei:\r\n\r\n„Voi lasa în mijlocul tau un popor smerit si mic, care se va încrede în Numele Domnului. Bucura-te si salta de veselie din toata inima ta, fiica Ierusalimului. Domnul Dumnezeul tau este în mijlocul tau ca un viteaz care poate ajuta; se va bucura de tine cu mare bucurie, va tacea în dragostea Lui, si nu va mai putea de veselie pentru tine. În ziua aceea va voi aduce înapoi; caci va voi face o pricina de slava si de lauda între toate popoarele pamîntului\" (tef. 3:12, 14, 17, 20).\r\n\r\nCuvinte cheia si teme caracteristice: tefania este cunoscut pentru vehementa cu care proclama „venirea zilei Domnului\" (tef. 1:14, 15; 2:3).\r\nMesajul cartii este clar: Dumnezeu nu poate fi înselat cu aparente trecatoare. El cunoaste si cerceteaza inima. Jumatatile de masura în pocainta sînt sinonime cu nepocainta. O jumatate de trezire spirituala este insuficienta pentru mîntuire tot asa cum este insuficienta o scara prea scurta sau un pod numai pe jumatate terminat.\r\nI. O PRIVIRE ÎNAUNTRU - MÎNIE ASUPRA LUI IUDA\r\na. Pedeapsa asupra lui Iuda, 1:1-18\r\nb. Îndemn la pocainta, 2:1-3\r\n\r\nII. O PRIVIRE ÎN JUR - MÎNIE ASUPRA NEAMURILOR\r\na. Vest si Est: Filistia si Moabul, 2:4-11\r\nb. Sud si Nord: Etiopia si Asiria, 4:12-15\r\nc. „Vai\"-uri asupra Ierusalimului, 3:1-8\r\n\r\nIII. O PRIVIRE DINCOLO DE TIMP - VREMEA TaMaDUIRII\r\na. Convertirea neamurilor, 3:9\r\nb. Restaurarea poporului Legamîntului, 3:10-15\r\nc. Slava se va întoarce în Ierusalim, 3:16-20" },
    //HAGAI
    { "HAGAI", "Explicatia Cartii: Hagai\r\n\r\nHagai, Zaharia si Maleahi sînt cei 3 profeti care au fost ridicati de Dumnezeu între evrei dupa revenirea din robia babiloniana. Cartile lor trebuiesc citite pe fondul evenimentelor istorice descrise de Ezra, Neemia si Estera. Cartea poarta numele celui ce a scris-o: „Hagai\". Acest nume este o prescurtare a lui „Hagaiah\" si se poate traduce prin „sarbatoarea Domnului\".\r\nNumele lui Hagai este mentionat de 9 ori în textul cartii (Hagai 1:1, 3, 12, 13; 2:1, 10, 13, 14, 20). Singura data cînd mai este mentionat în alta carte a Bibliei (Ezra 5:1; 6:14) îl întîlnim pe profet activînd alaturi de Zaharia încercînd sa trezeasca entuziasmului evreilor pentru rezidirea Templului.\r\nTimpul cînd a activat Hagai este foarte usor de aflat din datele pe care ni le furnizeaza textul la începutul fiecarui capitol. Profetul si-a rostit mesajele într-un interval de numai 4 luni, între l Septembrie si 24 Decembrie 520 î.Cr. Ne aflam la 15 ani dupa decretul dat de împaratul persan Cir, prin care s-a îngaduit evreilor sa se reîntoarca în tara lor si sa-si recladeasca Templul.\r\nÎn urma decretului dat de Cir (538 î.Cr.), o „ramasita\" din numarul total al evreilor raspînditi prin tarile imperiului s-au întors acasa. În convoiul condus de Zorobabel s-au înapoiat aproximativ 50.000 de iudei. Restul, estimat la cîteva milioane au preferat sa ramîna prin tinuturile în care îsi facusera un rost si o situatie, în tara s-au întors mai ales preotii, levitii si saracii. Entuziasmul reconstruirii pe ruinele trecutului glorios s-a spulberat însa repede sub povara greutatilor si sub amenintarile dusmanilor care nu vedeau cu ochi buni întarirea nationala si refacerea religioasa a poporului ales (Ezra 4-6). Gasind ca este mai usor sa astepte vremuri mai prielnice, decît sa riste o înfruntare directa cu dusmanii, evreii au stopat orice încercare de reconstructie a Templului în anui 534 î.Cr. Au trecut asa 14 ani de asteptare, încet, încet poporul si-a rezidit case pe temeliile vechilor asezari ale Ierusalimului. Templul iesise din preocuparile lor. Preocuparea pentru progresul economic fusese luata ca scuza pentru neglijarea vietii religioase de închinaciune. În acest context, Dumnezeu i-a ridicat pe Hagai si pe Zaharia ca sa mustre poporul si ca sa-i îndemne sa reia reconstruirea casei Domnului. Zaharia a început sa profeteasca între cel de al doilea si cel de al treilea mesaj rostit de Hagai. Sub impulsul îndemnurilor lor, poporul s-a apucat de treaba (520 î.Cr.) si reconstructia a fost terminata în anui 516 î.Cr. (Ezra 6:15). Desi pe afara totul parea în ordine, din Talmud aflam ca în aceasta noua cladire a Templului evreii nu au mai avut nici chivotul legamîntului si nici Urim si Tumim.\r\nMesajul profetic rostit de Hagai a constat într-o serie de 4 cuvîntari. Continutul lor a alternat între mustrare (primul si al treilea) si încurajare (al doilea si al patrulea). Poporul ajunsese mai preocupat cu înfrumusetarea caselor proprii decît de refacerea casei Domnului. Aceasta delasare în problemele religioase n-a fost usor de vindecat. Hagai a profetit numai 4 luni, dar Dumnezeu l-a ridicat apoi pe Zaharia care a profetit timp de 3 ani de zile. Dupa ei, Dumnezeu a continuat sa mustre si sa îndemne poporul prin Maleahi.\r\n\r\nMesajul proorocului Hagai este o chemare la o reasezare a prioritatilor în ierarhia stabilita de Dumnezeu. El avertizeaza poporul de pericolul neglijentei în slujirea Domnului. Hagai le arata evreilor ca atîta timp cît ei nu se vor angaja sa-l slujeasca pe Domnul, Dumnezeu le va încurca treburile si le va opri binecuvîntarea: „Uitati-va cu bagare de seama la caile voastre! Va asteptati la mult, si iata ca ati avut putin; l-ati adus acasa, dar Eu l-am suflat. Pentru ce? zice Domnul ostirilor. Din pricina Casei Mele, care sta darîmata, pe cînd fiecare din voi alearga pentru casa lui\" (Hagai l:7, 9).\r\n\r\nCuvinte cheie si teme caracteristice: Cartea lui Hagai marcheaza o piatra de hotar în istoria lui Israel. Rezidirea Templului nu s-a facut oricînd. Data fusese aleasa de Dumnezeu din vesnicie. Acesta a fost motivul pentru care, tocmai atunci, Domnul a trezit duhul lui Zorobabel prin lucrarea lui Hagai si a lui Zaharia. Se împlinisera cei 70 de ani de „pustiire\" rînduiti de Dumnezeu asupra Ierusalimului si asupra Templului. Ieremia anuntase ca: „tara aceasta va fi o paragina, un pustiu, si neamurile acestea vor fi supuse împaratului Babilonului timp de 70 de ani\" (Ier. 25:11). La încheierea acestei perioade, îngerul din cartea profetului Zaharia întreaba: „Doamne al ostirilor, pîna cînd nu vei avea mila de Ierusalim si de cetatile lui Iuda, pe care Te-ai mîniat în acesti sapte zeci de ani?\" (Zah. 1:12). Profetul Hagai este cel trimis sa anunte celor din Ierusalim raspunsul Domnului: „Din ziua aceasta, Îmi voi da binecuvîntarea Mea\" (Hagai 2:19).\r\nÎntreaga întîmplare este o ilustratie desavîrsita pentru afirmatia Domnului Isus Cristos din Matei 6:33: „Cautati mai întîi împaratia lui Dumnezeu si neprihanirea Lui, si toate celelalte lucruri vi se vor da pe deasupra\".\r\nI. ÎNDEMN LA REÎNCEPEREA REZIDIRII TEMPLULUI (1:1-15)\r\nExpresia caracteristica: „Suiti-va pe munte, aduceti lemne si ziditi Casa!\"\r\n\r\nII. ÎMBaRBaTARE ªI ÎNCURAJARE (2:1-9)\r\nExpresia caracteristica: „Fii tare... Caci Eu sînt cu voi si Duhul Meu este în mijlocul vostru\" (Hagai 2:4-5)\r\nIII. ÎNDEMNARE ªI PROCLAMARE (2:10-19)\r\nExpresia caracteristica: „Dar din ziua aceasta îmi voi da binecuvîntarea Mea\" (Hagai 2:19)\r\nIV. CONFIRMARE (2:20-23)\r\nExpresia caracteristica: „În ziua aceea, te voi lua si te voi pastra ca pe o pecete; caci Eu te-am ales\" (Hagai 2:23)" },
    //ZAHARIA
    { "ZAHARIA", "Explicatia Cartii: Zaharia\r\n\r\nDesi este mult mai scurta, cartea lui Zaharia poate fi asezata alaturi de profetiile lui Isaia, Daniel si Ezechiel în ceea ce priveste bogatia de informatii asupra „vremurilor viitoare\". Împreuna, aceste patru carti sînt corespondentul Apocalipsei din Noul Testament. Cartea poarta numele autorului ei. „Zekar-Iah\" se traduce prin: „Dumnezeu îsi aduce aminte\", nume cum nu se poate mai potrivit pentru cel ce anunta reînceperea vremurilor de binecuvîntare asupra Israelului.\r\nZaharia este un nume dat copiilor nascuti la batrînete. În textul Scripturii ne întîlnim cu nu mai putin de 29 de personaje care au purtat aceasta numire.\r\n\r\nProfetul Zaharia a fost „fiul lui Berechia, fiul lui Ido\" (Zah. 1:1). Ca si predecesorii sai, Ieremia si Ezechiel Zaharia a apartinut neamului preotesc (Ezra 51; 6:14; Neemia 12:4, 16). El a fost nascut în Babilon si cînd împaratul Cir a dat decretul reîntoarcerii, a venit împreuna cu bunicul sau înapoi în Palestina. A fost chemat de Dumnezeu la slujba profetiei înca de tînar (Zah. 2:4) si a activat la început împreuna cu mai vîrstnicul Hagai (Ezra 5:l; 6:14). Lucrarea lui s-a întins pe o perioada de 3 ani de zile. Domnul Isus ne spune în Matei 23:35 ca Zaharia a fost „omorît între Templu si altar\", sfîrsind la fel cu un alt purtator al aceluiasi nume (2 Cron. 24:20-21).\r\nÎmpreuna cu Zorobabel, capetenia poporului, Iosua, marele preot si Hagai profetul, Zaharia a fost folosit de Dumnezeu în Ierusalim începînd cu anul 520 î.Cr.\r\nCa si Hagai, Zaharia este un profet al reconstructiei. El este ridicat de Dumnezeu sa vorbeasca evreilor care se întorsesera în Palestina dupa cei 70 de ani petrecuti în robia Babiloniana.\r\n\r\nAltadata un neam puternic si vestit, evreii ajunsesera sa locuiasca în tara lor doar datorita bunavointei unui stapînitor al lumii de atunci. Situtia lor materiala, politica si sociala era vrednica de toata mila. În acest context, Dumnezeu i-a folosit pe Hagai si Zaharia sa le spuna celor întorsi din Babilon ca nu trebuie sa se opreasca la prezentul dezolant, ci trebuie sa spere si sa lucreze pentru viitorul glorios pe care li l-a pregatit Iehova. Hagai a avut un mesaj aspru si critic; prin Zaharia, Dumnezeu a lucrat altfel: mesajul lui Zaharia a trecut repede peste neajunsurile prezentului descoperindu-le viitorul stralucit care îi asteapta. Zaharia este un vestitor al slavei si al lucrarilor lui Mesia.\r\nPentru a descrie planurile viitoare pe care le-a facut Dumnezeu cu poporul ales, Zaharia â folosit o sene de opt vedenii, patru predici si doua rostiri profetice. Primele 8 capitole ale cartii au fost scrise ca o încurajare a poporului în lucrarea de rezidire a Templului. Ultimele 6 capitole au fost adaugate dupa terminarea reconstructiei si prevestesc lucrarea lui Dumnezeu prin Mesia. Mesajul din cartea lui Zaharia a trecut de la stapînirea popoarelor straine la stapînirea lui Mesia, de la pedeapsa la vremuri de pace si de la nelegiure la trairea în sjîntenia prezentei Domnului.\r\n\r\nCuvinte cheie si teme caracteristice: în cartea lui Zaharia exista o extraordinar de amanuntita prezentare a lucrarii lui Mesia:\r\n\r\nCristos, „Odrasla\" - Zah. 3:8\r\nCristos, „Robul Domnului\" - Zah. 3:8\r\nIntrarea în Ierusalim calare pe un magarus - Zah. 9:9\r\nCristos, „Pastorul cel bun\" - Zah. 9:16; 11:11\r\nCristos, pastorul lovit cu oile risipite - Zah. 13:7\r\nCristos, vîndut pe 30 de arginti, 11:12-13\r\nCristos cu mîinile strapunse - Zah. 12:10\r\nCristos, Mîntuitorul poporului Sau - Zah. 12:10; 13:1\r\nCristos, ranit în casa celor ce-L iubeau - Zah. 13:6\r\nCristos, se întoarce pe muntele Maslinilor - Zah. 14:3-8\r\nCristos, încoronat la venirea Lui - Zah. 14\r\n\r\nMesajul peste veacuri: În afara de confirmarea lui Isus din Nazaret ca Mesia prin împlinirea profetiilor sale, Zaharia ne ofera si o lectie de spiritualitate în textul din capitolul 7. O delegatie venita din partea oamenilor din Betel îl întrebasera daca trebuiesc reluate ritualurile sarbatoresti de altadata. Raspunsul dat atunci de Zaharia este o lectie care nu si-a pierdut nici astazi valabilitatea: Sarbatorile fara suportul unei vieti de ascultare fata de Domnul sînt fara nici o valoare înaintea ochilor Lui.\r\nI. PROFETII DE ÎNCEPUT 1 - 8\r\na. Chemare la pocainta, 1:1-6\r\nb. Vedenia calaretilor, 1:7-17\r\nc. Vedenia celor patru coarne, 1:18-21\r\nd. Vedenia celui cu cumpana, 2:1-13\r\ne. Vedenia marelui preot, 3:1-10\r\nf. Vedenia sfesnicului de aur, 4:1-14\r\ng. Vedenia sulului de carte zburator, 5:1-4\r\nh. Vedenia femeii cu efa, 5:5-11\r\ni. Vedenia celor 4 care, 6:1-8\r\nî. Încoronarea lui Iosua, 6:9-15\r\nj. Posturi si binecuvîntari 7:1 - 8:23\r\n\r\nII. PROFEtII TÎRZII 9-14\r\na. Sosirea Împaratutui-Pastor, 9-10\r\nb. Ranirea si uciderea Împaratui-Pastor, 11\r\nc. Biruinta finala si încoronarea în slava, 12-14" },
    //MALHEAI
    { "MALEAHI", "Explicatia Cartii: Maleahi\r\n\r\nCartea lui Maleahi este ultima carte a Vechiului Testament. Maleahi este ultimul glas care rasuna în Israel înainte ca vocea profetilor sa amuteasca pentru cîteva sute de ani. Cartea poarta numele autorului ei. „Maleahi \"este un derivat de la „Mal-ac-Iah\"care se poate traduce prin „Mesagerul lui Iehova\".\r\nProfetul Maleahi a fost ridicat de Dumnezeu în Israel la putin timp dupa activitatea desfasurata în Ierusalim de Neemia. Personalitatea si familia profetului sînt complet necunoscute. Asa cum îl arata si numele, profetul este important numai pentru mesajul pe care-l poarta, nu si prin sine însusi.\r\nCu siguranta ca ne aflam dupa revenirea evreilor în Ierusalimul lui Iuda, la suficient timp dupa activitatea lui Neemia pentru ca poporul sa o fi luat iarasi pe cai gresite. Zelul preotilor si al poporului se stinsese. Activitatea religioasa cazuse iarasi prada formalismului (Mal. 3:14) si acesta incomplet si neconform Legii (Mal l: 14). Ultima data cunoscuta despre activitatea lui Neemia îri Ierusalim a fost anul 430 î.Cr. asa ca trebuie sa plasam misiunea lui Maleahi la putin timp dupa aceasta data, probabil între anii 420- 397 î.Cr.\r\nData la care Dumnezeu l-a ridicat pe Maleahi în Israel este legata direct de profetia rostita de Daniel asupra Ierusalimului. Daca va aduceti aminte, profetului i se spusese ca asupra Ierusalimului vor trece pîna la vremea sfîrsitului 70 de saptamîni de ani (Dan. 9:24). Exprimarea lui a parut putin neclara si inutil complicata prin felul în care a facut el enumerarea acestor saptamîni: „Vor trece sapte saptamîni; apoi... sasezeci si doua de saptamîni\" (Dan. 9:25). Ce a fost asa de important dupa primele sapte saptamîni pentru ca proorocul sa marcheze acea „borna\" din scurgerea vremii? Este greu sa gasim un alt raspuns mai bun decît faptul ca la „sapte saptamîni\" de ani în Israel a încetat activitatea profetica. Mesajul rostit de Maleahi a fost evenimentul care a marcat scurgerea primelor sapte saptamîni de ani din calendarul profetic hotarît de Dumnezeu asupra Israelului.\r\n\r\nProblemele din vremea lui Maleahi seamana foarte mult cu cele întîlnite de Neemia în activitatea lui: o preotie corupta (Mal. 1:6; Neemia 13:1-9), neglijarea darniciei catre casa Domnului (Mal. 3:7; Neemia 13:10-13) si casatorii încheiate cu femei din afara Israelului (Mal. 2:10-16; Neemia 13:23-28).\r\nScrisa în forma clara a unui dialog între Dumnezeu si evrei, cartea lui Maleahi este un „apel\" insistent la pocainta si la întoarcerea la Dumnezeu. Profetia se împarte simetric în doua parti: în primele doua capitole apelul la pocainta este motivat de starea prezenta a poporului, iar în ultimele doua capitole chemarea la pocainta este motivata de venirea „zilei Domnului\".\r\n\r\nCuvinte cheie si teme caracteristice: Întreaga cartea este un manifest împotriva îndepartarii de Dumnezeu. Textul ei poate fi folosit foarte bine pentru a-i trezi pe cei care s-au „racit\" în sentimentele lor fata de Domnul. Punctul culminant al profetiei este Maleahi 3:1, 2: „Iata voi trimite pe solul Meu; el va pregati calea înaintea Mea. ªi deodata va intra în Templul Sau Domnul pe care-L cautati: Solul legamîntului pe care-L doriti;... Cine va putea sa sufere însa ziua venirii Lui? Cine va ramînea în picioare cînd Se va arata El?\"\r\n\r\nMaleahi întinde mîna peste aproximativ patru sute de ani si leaga Vechiul Testament de Noul Testament prin vestirea lucrarilor lui Ioan Botezatorul si ale Domnului Isus.\r\n\r\nNumirea de „mesager al lui Iehova\" i s-a potrivit de minune lui Maleahi, dar tot la fel de bine prin „sol al Domnului\" poate fi subînteles Ioan Botezatorul si chiar Domnul Isus.\r\n\r\nCuvînt de încheiere: Se cuvine sa subliniem cîteva lucruri care-l fac pe Maleahi: omul care închide canonul Vechiului Testament:\r\n\r\n1. Cartea lui se încheie cu un cuvînt semnificativ pentru starea în care a ajuns poporul sub administratia Legii Mozaice: „blestem\" (Mal. 4:6). Neascultarea evreilor a atras asupra lor dezlantuirea pedepsei. Trebuia ca Cineva sa vina si sa înlature acest blestem care apasa greu asupra poporului.\r\n\r\n2. Profetia lui Maleahi contine anuntarea venirii Domnului Isus (Mal. 3:1-3). El este Cel ce „s-a facut blestem pentru noi\" lasîndu-se sa fie atîrnat în locul nostru pe lemnul Crucii.\r\n\r\n3. Maleahi publica un avertisment pentru toti cei ce asteapta ziua Domnului. Exista pericolul asteptarii Domnului într-o atitudine gresita. Scepticismul, indiferenta si formalismul sînt semnele ei caracteristice.\r\n\r\n4. Maleahi anunta preocuparea lui Dumnezeu pentru cei care nu-L uita (Mal. 3:15-18). „O carte de aducere aminte a fost scrisa înaintea Lui, pentru cei ce se tem de Domnul, si cinstesc Numele Lui\".\r\n\r\nCartea lui Maleahi se încheie într-o atmosfera de neglijenta si necinstire a Numelui lui Dumnezeu. Doar membrii unei „ramasite\" mai vorbesc unul cu altul soptind încet: „ªtii ca El vine?\" „Da, stiu!\" Dumnezeu nu-i uita pe acestia si-i noteaza pentru binecuvîntarea lor vesnica.\r\nI. PRIMUL APEL 1-2\r\na. Grija parinteasca, 1:1-5\r\nb. Înselaciunea, 1:6-14\r\nc. Necredinciosia, 2:1-9\r\nd. Casatorii nepermise, 2:10-12\r\ne. Divorturi, 2:13-16\r\nf. Obraznicie, 2:17\r\n\r\nII. AL DOILEA APEL 3 – 4\r\na. Venirea Solului Domnului, 3:1-6\r\nb. Hotie, 3:7-12\r\nc. Aroganta, 3:13-15\r\nd. „Ramasita credincioasa\", 3:16-18\r\ne. Ziua Domnului, 4:1-6" },

     // Noul Testament
     //MATEI
    { "MATEI", "Explicatia Cartii: Matei\r\n\r\nTitlul: În originalul grec, evanghelia poartă titlul: \"Kata Mataion\"-\"după Matei.\"\r\n\r\nAutorul: Experienţa convertirii lui Matei ne este dată de el însuşi (Mt 9.9-17). Numele lui cel vechi, \"Levi, fiul lui Alfeu\" (Mc 2.14), este înlocuit cu \"Matei\" care tălmăcit înseamnă \"darul lui Dumnezeu.\" Probabil că această poreclă i-a fost dată ca o amintire a harului pe care i l-a făcut Domnul Isus atunci când l-a chemat la Sine. Aduceţi-vă aminte că Matei şi-a deschis inima şi casa ca să-L găzduiască pe Domnul Isus. Nu numai atât, dar, aşa cum a remarcat un comentator, Matei şi-a luat cu sine şi până atunci când a părăsit vama ca să-L urmeze pe Domnul. Mâinile lui care serviseră altădată înrobirea şi jaful, s-au dat în întregime lui Isus.\r\n\r\n\r\nConţinutul cărţii: Nici una dintre cele 4 Evanghelii nu este o biografie în sensul modern al cuvântului. De fapt, apostolul Ioan îşi exprimă îndoiala că o astfel de biografie ar fi posibil de întocmit vreodată:\r\n\"Mai sunt multe alte lucruri, pe care le-a făcut Isus, care, dacă s-ar fi scris cu deamăruntul, cred că nici chiar în lumea aceasta n-ar fi putut încăpea cărţile care s-ar fi scris.\" (In 21.25).\r\nMatei îl prezintă pe Domnul Isus drept împăratul mesianic promis de Dumnezeu Israelului (Mt 1.23; Mt 2.2,6; Mt 3.17; Mt 4.15-17; Mt 21.5,9; Mt 22.44,45; Mt 26.64; Mt 27.11,27-37). Fiind obişnuit să ţină evidente sistematice, fostul vameş ne dă o cronică foarte ordonată a vieţii şi misiunii împărăteşti a Domnului Isus. Despre venirea împărăţiei a vorbit Ioan Botezătorul (Mt 3.1-2). Acelaşi mesaj l-a avut şi Domnul Isus în debutul misiunii Sale publice (Mt 4.23). Aceiaşi proclamaţie au fost trimişi să o facă şi cei 12 apostoli ai Domnului (Mt 10.1-7).\r\n\r\n\r\nCuvinte cheie şi teme caracteristice: Numirea \"împărăţia cerurilor\" apare de 28 de ori în textul evangheliei lui Matei şi niciodată în celelalte scrieri ale Noului Testament. Evanghelistul foloseşte 130 de citate din cărţile Vechiului Testament pentru a le dovedi evreilor că Isus Cristos este Mesia. Expresia: \"ca să se împlinească ceea ce fusese vestit prin prooroci\" este folosită de 9 ori de Matei şi nici ea nu mai apare în textul vreuneia dintre celelalte Evanghelii.\r\n\r\nEvanghelia lui Matei este singura în care este menţionat cuvântul: \"biserică\" (Mt 16.18; Mt 18.17). Termenul grecesc folosit \"e????s?a\" - eclesia - înseamnă \"adunarea celor chemaţi afară cu un scop\", \"mulţimea celor puşi de o parte.\" În Vechiul Testament, Israel fusese \"poporul ales al lui Dumnezeu.\" Situaţia aceasta începuse odată cu chemarea lui Avraam şi cu punerea lui deoparte (Gen 12.1; Deut 7.6-8). De fapt, Ştefan numeşte naţiunea lui Israel chiar cu termenul de \"Biserică\" atunci când spune: \"El este acela care, în adunarea Israeliţilor din pustie...\" (Fapte 7:38). Evreii fuseseră poporul \"pus deoparte de Dumnezeu\", în Noul Testament, Biserica nu mai este compusă numai din evrei. În această adunare nu se mai fac distincţii rasiale (Gal 3.28). Din ea pot face parte şi evreii şi cei dintre neamuri. Chiar şi Matei, care scrie în mod special evreilor, are în Evanghelia lui elemente \"universaliste\" care îi includ şi pe cei dintre Neamuri. Iată numai câteva exemple: Magi din răsărit vin să se închine pruncului Isus (Mt 2.1-12), Domnul face minuni pentru cei dintre Neamuri şi chiar îi laudă pentru credinţa lor (Mt 8.5-13; Mt 15.21-28), împărăteasa din Seba este lăudată pentru osteneala pe care a depus-o ca să vină să audă înţelepciunea dumnezeiască aşezată în Solomon (Mt 12.42), în ceasul hotărâtor al misiunii Sale, Domnul Isus se întoarce la o profeţie făcută despre Neamuri (Mt 12.14-21), chiar şi în pildele pe care le dă, Isus arată că binecuvântările refuzate de Israel vor fi împărţite Neamurilor (Mt 22.8-10; Mt 21.40-46), predica rostită pe Muntele Măslinilor promite că mesajul Evangheliei \"va fi propovăduit în toată lumea , ca să slujească drept mărturie tuturor neamurilor\" (Mt 24.14), iar împuternicirea dată ucenicilor după înviere este de a se duce să facă ucenici \"din toate neamurile.\" (Mt 28.19-20).\r\n\r\n\r\nCUPRINSUL CĂRŢII\r\n\r\n\r\nI. APARIŢIA ÎMPĂRATULUI Mt 1-10\r\n\r\n\r\na. Persoana Sa (Mt 1.0-4.0);\r\nb. Principiile Sale (Mt 5.0-.07);\r\nc. Puterea Sa (Mt 8.0-10.0).\r\n(Notă: Anunţul caracteristic acestei perioade este: \"împărăţia cerurilor este aproape\" (Mt 3.2; Mt 4.17; Mt 10.7)\r\n\r\n\r\nII. RĂSCOALĂ ÎMPOTRIVA LUI Mt 11-13\r\n\r\n\r\na. Mesajul Său este respins (Mt 11.1-9);\r\nb. Lucrarea Lui este negată (Mt 11.20-30);\r\nc. Principiile Lui sunt respinse (Mt 12.1-21);\r\nd. Persoana Lui este atacată (Mt 12.22-50);\r\ne. Urmarea: împărăţia devine o \"taină\" refuzată acelei generaţii (pildele împărăţiei din Mt 13.0).\r\n\r\n\r\nIII. RETRAGEREA ÎMPĂRATULUI Mt 14-20\r\n\r\n\r\n(Domnul părăseşte mulţimile pentru a petrece timpul împreună cu grupul restrâns al ucenicilor);\r\n\r\n\r\na. înainte de mărturisirea lui Petru (Mt 14.1-16.12);\r\nb. Mărturisirea lui Petru (Mt 16.13-28);\r\n\r\n\r\n(Cea dintâi anunţare a Crucii - Mt 16.21).\r\n\r\n\r\nc. După mărturisirea lui Petru (Mt 17.1-20:34);\r\n\r\n\r\n(A doua anunţare a Crucii - Mt 17.22);\r\n(A treia anunţare a Crucii - Mt 20.17-19).\r\n\r\n\r\nIV. LEPĂDAREA ÎMPĂRATULUI Mt 21-27\r\n\r\n\r\n(\"De aceea, vă spun că împărăţia lui Dumnezeu va fi luată de la voi,...\" Mt 21.43)\r\n\r\n\r\na. Prezentarea Lui însuşi ca împărat (Mt 21.1-16);\r\nb. înfruntarea cu conducătorii (Mt 21.17-23.39);\r\nc. Mesajul Său profetic (Mt 24-25);\r\nd. Suferinţele şi moartea Sa (Mt 26-27).\r\n\r\n\r\nV. ÎNVIEREA ÎMPĂRATULUI Mt 28.0" },
    //MARCU
    { "MARCU", "Explicatia Cartii: Marcu\r\n\r\nTitlul: În originalul grec, cartea poartă titlul: \"Kata Markon\" - \"după Marcu.\"\r\n\r\nAutorul: Să trecem în revistă mai întâi ceea ce ştim despre acest personaj:\r\n\r\na. Mama lui s-a numit Maria (Miriam - în ebraică), dar el însuşi a purtat un nume evreiesc \"Ioan\" unul roman \"Marcus\", ceea ce ne poate arăta sau că a avut un tată roman sau că tatăl său îşi câştigase (sau cumpărase, vezi Fapt 22.28) cetăţenia romană.\r\nb. Familia lui s-a bucurat de o oarecare prosperitate materială, subliniată de altfel şi de prosperitatea unchiului său Barnaba (Fapt 4.37).\r\nc. În casa lor s-au întâlnit adeseori ucenicii Domnului Isus după înălţarea învăţătorului lor la cer. De fapt, Ioan Marcu a fost prins fără voia lui în lanţul evenimentelor şi se pare că era cât pe ce să o păţească pentru amestecarea lui în anturajul care-l însoţea adesea pe \"Rabinul din Nazaret.\" Cu siguranţă că el este acela despre care se vorbeşte în Mc 14.51-52.\r\n\"După El mergea un tânăr, care n-avea pe el decât o învelitoare de pânză de in. Au pus mâna pe el; dar el şi-a lăsat învelitoarea, şi a fugit în pielea goală.\"\r\nNumai Ioan Marcu putea cunoaşte şi scrie asemenea detalii despre propria lui păţanie! d. Despre Marcu mai aflăm şi că a fost \"fiul lui Petru\":\r\n\"Biserica aleasă cu voi, care este în Babilon, vă trimete sănătate. Tot aşa şi Marcu, fiul meu.\" (1Pet 5.13). Această exprimare plină de afecţiune ne arată două lucruri: că Petru l-a \"născut\" pe Marcu în credinţa creştină şi că de-a lungul întregii sale vieţi, Marcu a continuat să fie un fiu al lui Petru. Există evidenţe clare că între cei doi au fost legături comune de slujire.\r\n\r\nPapias, unul din presbiterii Bisericii din secolul întâi, scrie că Ioan apostolul a spus următoarele. \"Marcu, fiind traducătorul lui Petru, a notat totul cu fidelitate, nu într-o ordine cronologică a evenimentelor, căci el n-a umblat cu Domnul, ci doar l-a auzit pe Petru povestind.\" Se pare că Petru, pescar lipsit de o educaţie aleasă, a continuat să vorbească cu predilecţie limba aramaică şi că de câte ori a trebuit să vorbească în afara Iudeii l-a folosit pe Ioan Marcu în slujba de interpret. Ani de-a rândul deci, Ioan Marcu i-a fost \"gura\" prin care Petru a vorbit Neamurilor. Probabil că după moartea marelui apostol, oamenii l-au îndemnat pe Ioan Marcu să continue să le povestească ceea ce spusese Petru. Spre bătrâneţe, Marcu însuşi a aşternut pe hârtie viaţa Domnului Isus aşa cum o ştia de la Petru şi această scriere a lui a devenit cea de a doua Evanghelie aşezată în Noul Testament. Cine o va citi ştiind toate acestea, va recunoaşte uşor pasajele în care este evidentă contribuţia lui Petru şi îşi va explica de ce această scriere \"din memorie\" este mai scurtă decât toate celelalte trei Evanghelii.\r\n\r\nConţinutul cărţii: Deşi a ştiut multe despre Domnul Isus, Marcu a scris cea mai scurtă dintre cele patru evanghelii. Toţi cei ce au studiat evanghelia lui Marcu au remarcat că ea este o carte de acţiune. Marcu nu este preocupat nici de împlinirea profeţiilor şi nici de clarificarea genealogiilor. Discursurile, atunci când apar, sunt date în forme prescurtate. Nu întâlnim în textul lui Marcu nici un fel de genealogie a Mântuitorului. Accentul este pus pe activitatea depusă de Domnul Isus în slujba oamenilor.\r\n\r\nCuvinte cheie şi teme caracteristice: Evanghelia lui Marcu nu are nici un fel de introducere. Toate celelalte au un cuvânt de lămurire sau un preambul care să prezinte scopul scrierii. Marcu trece direct la redarea faptelor vieţii Domnului Isus. Cuvântul caracteristic lui Marcu este: \"îndată...\" El este repetat mereu, de parcă autorul ar fi dorit să ne atragă atenţia că personajul pe care ni-L prezintă este mereu grăbit şi ocupat cu lucrarea pe care a venit să o facă. Marcu nu redă discursurile Domnului Isus, dar ne dă în schimb minunile săvârşite de El. În textul evangheliei sale găsim nu mai puţin de 20 de minuni săvârşite de Domnul.\r\nTema evangheliei lui Marcu este: \"Cristos - Robul\" şi se găseşte în textul din Mc 10.45: \"Căci Fiul omului n-a venit ca să I se slujească, ci El să slujească, şi să-Şi dea viaţa răscumpărare pentru mulţi.\"\r\n\r\nCUPRINSUL CĂRŢII\r\n\r\nI. Slujirea lui Cristos ca Rob, Mc 1-10\r\n\r\na. Pregătirea, Mc 1.1-13:\r\n\r\n1. prin misiunea lui Ioan Botezătorul, Mc 1.1-1;\r\n2. prin botezul Lui, Mc 1.9-11;\r\n3. prin ispitirile Lui, Mc 1.12-13.\r\n\r\nb. Propovăduirea Lui, Mc 1.14-20;\r\nc. Puterea Lui, Mc 1.21-3.12;\r\nd. Anturajul Lui, Mc 3.13-35;\r\ne. Pildele Lui, Mc 4.1-34;\r\nf. Prerogativele Lui, Mc 4.35-9.1;\r\ng. Prevestirile Lui, Mc 9.2-50;\r\nh. Predicarea în Berea, Mc 10.1-52.\r\n\r\nII. Sacrificiul lui Cristos ca Rob, Mc 11-15\r\n\r\na. Duminică: Intrarea triumfală în Ierusalim, Mc 11.1-11;\r\nb. Luni: Blestemarea smochinului si curăţirea Templului, Mc 11.12-19;\r\nc. Marţi: Discursuri şi învăţături, Mc 11.20-13.37;\r\nd. Miercuri: Uns de Maria şi trădat de Iuda, Mc 14.1-11;\r\ne. Joi: Cina şi trădarea, Mc 14.12-52;\r\nf. Vineri: Judecata şi pătimirea, Mc 14.53-15.47.\r\n\r\nIII. Biruinţa lui Cristos ca Rob, Mc 16\r\n\r\na. Duminică: Învierea Lui, Mc 16.1-8\r\nb. Arătările Lui, Mc 16.9-18\r\nc. Înălţarea Lui, Mc 16.19-20" },
    //LUCA
    { "LUCA", "Explicatia Cartii: Luca\r\n\r\nTitlul: În originalul grec, cartea poartă titlul: „Kala Loukon\" - „după Luca\".\r\n\r\nAutorul: Din introducerile pe care le găsim la începutul „Evangheliei lui Luca\" şi a cărţii „Faptele Apostolilor\" se vede foarte clar că ambele au fost scrise de un singur autor: „Luca\" şi că amîndouă au fost destinate aceluiaşi om: \"Teofil\". Cine a fost acest Luca?\r\n\r\n1. Un tovarăş de călătorie al lui Pavel. Schimbarea de la „au trecut atunci prin Misia\" din Fapte 16:8, la „După vedenia aceasta a lui Pavel, am căutat îndată să ne ducem în Macedonia...” din 16:10 ne spune că Luca s-a alăturat apostolului Pavel la Troa. După aceasta, Luca a devenit tovarăşul de lucru nedespărţit al lui Pavel. Cu el a călătorit la Filipi (16:12), iar şase ani mai târziu tot cu el a plecat din Filipi (20:6). Au fost împreună la Ierusalim (21:17), cînd mulţimea dezlănţuită a încercat să-l linşeze pe Pavel. Au fost împreună în cei doi ani de detenţie la Cezareea (24:27 şi 27:1). Împreună au fost şi în călătoria plină de peripeţii înspre Roma, cînd s-a spart corabia cu ei (27:1 - 28:16). Împreună au fost la Roma, cu ocazia procesului judecat de Nero şi tot împreună au fost, se pare, şi în ultimele clipe dinainte de martirajul marelui apostol (Col. 4:14; 2 Tim.4:2; Filimon 24).\r\n\r\n2. Medic iubit de apostoli. Aflăm lucrul acesta din încheierea scrisorii lui Pavel către Biserica din Colose: „Luca, doctorul prea iubit...” (Col. 4:14). Probabil că Luca a dat îngrijiri medicale multora şi în multe locuri, ajungînd să fie iubit şi preţuit de Biserici. Dacă citim bine, vedem că şi însoţirea cu Pavel s-a făcut într-o vreme cînd Pavel avea nevoie de îngrijire medicală datorită unui atac acut de boală a ochilor (compară „Galatia\" din Fapte 16:6 cu „noi\" din 16:10 şi Galateni 4:13-15)\r\n\r\nConţinutul cărţii: Evanghelia lui Luca poate fi numită Evanghelia Omului Isus Cristos. Această afirmaţie nu este o blasfemie care neagă dumnezeirea Celui născut în Nazaret, ci este o proclamare a „întrupării\" care a făcut posibilă „ispăşirea şi răscumpărarea\". Luca ni-L prezintă pe Domnul Isus ca descendent, nu numai al lui Avraam (lucrarea lui Matei), ci şi al lui ADAM, strămoşul tuturor oamenilor (vezi genealogia din cap. 3:23-38). Prin aceasta el îl aşează în şuvoiul scurgerii istoriei umane, ca Cel venit nu numai să participe deplin, dar şi să reaşeze omenirea în prerogativele pierdute de neascultarea lui Adam. Învierea Domnului Isus dintre cei morţi este dovada dată tuturor de Dumnezeu că „în Isus\" se poate reveni „acasă\". Viaţa veşnică este oferită încă o dată omenirii prin Golgota. Cristos a venit să se facă pentru un timp „ca noi\" pentru a putea să ne facă pentru veşnicie „ca El\".\r\n\r\nCuvinte cheie şi teme caracteristice: Faptul că a stat atît de mult alături de Pavel îl face pe Luca să fie legătura firească dintre Dumnezeul evreilor, prin apostolul neamurilor, către oamenii de pretutindeni. Influenţa exercitată de propovăduirea lui Pavel este evidentă în teologia Evangheliei scrisă de Luca. Remarcaţi cum la capitolul: „Instituirea Cinei Domnului\" asemănările merg chiar şi pînă la folosirea aceloraşi cuvinte.\r\n\r\nCUPRINSUL CĂRŢII\r\nPrefaţa: Metoda şi scopul scrierii, 1:1-4A.\r\n\r\nI. Identificarea Fiului Omului, 1:5-4:13\r\nA. Anunţarea naşterii lui Ioan Botezătorul, 1:5-25\r\nB. Anunţarea naşterii Fiului Omului, 1:26-56\r\nC. Naşterea lui Ioan Botezătorul, 1:57-80\r\nD. Naşterea Fiului Omului, 2:1-20\r\nE. Isus ca prunc adus la Templu, 2:21-38\r\nF. Isus ca şi copil adus la Templu, 2:39-52\r\nG. Botezul Fiului Omului, 3:1-22\r\nH. Genealogia Fiului Omului, 4:1-13\r\n\r\nII. Slujirea Fiului Omului, 4:14-9:50\r\nA. Începuturile slujirii Lui, 4:14-30\r\nB. Autoritatea slujirii Lui, 4:31-6:11\r\n\r\nC. Tovarăşii Săi de lucru, 6:12-49\r\n1. Chemarea ucenicilor, 6:12-16\r\n2. Caracteristicile ucenicilor, (Marea Predică), 6:17-49\r\n\r\nD. Lucrările slujirii Lui, 7:1-9:50\r\n1. slujirea celor bolnavi, 7:1-10\r\n2. slujirea celor morţi, 7:11-17\r\n3. slujirea celor îndoielnici, 7:18-35\r\n4. slujirea celor păcătoşi, 7:36-50\r\n5. alte slujiri, 8:1 -9:50\r\n\r\nIII. Lepădarea Fiului Omului, 9:51-19:27\r\nA. Refuzat de samariteni, 9:51-56\r\nB. Refuzat de oamenii din preajmă, 9:57-62\r\nC. Trimiterea celor 70, 10:1-24\r\nD. Refuzat de un învăţător al Legii, 10:25-37\r\nE. Primit în Betania, 10:38-42\r\nF. Despre rugăciune, 11:1-13\r\nG. Refuzat de neam, 11:14-36\r\nH. Refuzat de farisei şi cărturari, 11:37-54\r\nI. Refuzat, dar îi mai învaţă o dată, 12:1-19:27\r\n\r\nIV. Osîndirea Fiul Omului, 19:28-23:56\r\nA. Duminică: 19:28-44\r\nB. Luni: 19:45-48\r\nC. Marţi: 20:1-21:38\r\nD. Miercuri: 22:1-6\r\nE. Joi: 22:7-53\r\nF. Vineri: 22:54-23:55\r\nG. Sîmbătă: 23:56\r\n\r\nV. Biruinţa Fiului Omului, 24:1-53\r\nA. Biruitor asupra morţii, 24:1-12\r\nB. Împlinitor al profeţiilor, 24:13-35\r\nC. Biruitor asupra îndoielilor ucenicilor, 24:36-43\r\nD. Biruitor dăruit Bisericii, 24:44-48\r\nE. Biruitor împărţind biruinţa, 24:49\r\nF. Biruitor înălţat în slava, 24:50-53" },
    //IOAN
    { "IOAN", "Explicatia Cartii: Ioan\r\n\r\nTitlul: În originalul grec, cartea poartă titlul: „Kata Ioannen\" - „după Ioan\".\r\n\r\nAutorul: Despre Ioan ştim că a fost fiul lui Zebedei şi frate cu Iacov, împreună se ocupau cu pescuitul şi s-au întîlnit pentru prima dată cu Domnul Isus pe malul mării, pe cînd îşi cîrpeau mrejele (Matei 4:21; Marcu 1:19). Ioan a avut toată viaţa lui o relaţie deosebită cu Simon Petru, pe care l-a cunoscut încă dinainte de a deveni ucenic al Domnului Isus. Luca 5:10 ne spune că ei pescuiau împreună. Probabil că Petru era încă de pe atunci un om pe care vîrsta şi mai ales personalitatea l-au ajutat să ocupe o poziţie de lider între ceilalţi. Marcu ne spune că aceşti pescari aveau deja creiată între ei o oarecare intimitate:\r\n\r\n„După ce a ieşit din sinagogă, a intrat împreună cu Iacov şi Ioan în casa lui Simon şi a lui Andrei\" (Marcu 1:29). Pe Ioan şi Petru îi întîlnim apoi împreună în grupul restrîns de ucenici pe care i-a luat Domnul Isus cu Sine peste tot (Petru, Iacov şi Ioan sînt împreună cu Isus pe Muntele schimbării la faţă, în grădina Gheţimani, etc.). Tot împreună sînt trimişi să pregătească Pastele înainte de răstignirea Domnului (Luca 22:8), împreună se duc în curtea Marelui Preot (Ioan 18:16), împreună îi vedem fugind la mormînt să verifice învierea (Ioan 20:1-10) şi tot împreună îi găsim după înălţarea lui Isus la cer, în activităţile Bisericii primare (Fapte 3:1). Cei doi fuseseră tovarăşi în meseria de pescari şi Mîntuitorul i-a transformat în „pescari de oameni\" (Marcu 1:17).\r\n\r\nÎnainte de a deveni ucenic al lui Isus, Ioan a fost ucenic al lui Ioan Botezătorul (Ioan l:35-40).\r\n\r\nConţinutul cărţii: Spre deosebire de ceilalţi trei Evanghelişti care s-au străduit să aştearnă pe hîrtie viaţa Domnului Isus într-o ordine oarecum cronologică, Ioan porneşte în scrierea Evangheliei sale de la cu totul alte premize.\r\n\r\nEvanghelia lui este gîndită ca o necesară „completare\" a celorlalte trei. Cu siguranţă că aţi remarcat faptul acesta. Primii Evanghelişti ne-au lăsat scris de unde a venit Isus, în ce sătuc a fost născut, care i-au fost părinţii, copilăria, peregrinările, peripeţiile, lucrările, moartea şi învierea, dar, concentrîndu-se preponderent asupra a ceea ce a făcut Domnul Isus, ei nu ne-au spus destul despre cine a fost de fapt El. Supravieţuind ca longevitate tuturor celorlalţi apostoli, şi ajungînd să trăiască într-o epocă în care divinitatea, nu istoricitatea lui Cristos era contestată, Ioan s-a apucat să scrie căutînd să-şi convingă cititorii că: „Isus este\" nu numai fiul Mariei şi al lui Iosif, ci şi „Cristosul, Fiul lui Dumnezeu\". Numai un astfel de Isus a putut aduce oamenilor mîntuirea: „şi crezînd, să aveţi viaţa în Numele Lui\" (Ioan 20:31)\r\n\r\nAcesta este punctul central în care Ioan se deosebeşte de ceilalţi Evanghelişti. Cei trei îl prezintă pe Isus, el merge mai departe şi-L interpretează pe Isus. Ei ni-L prezintă pe Domnul din afară, Ioan pătrunde mai adînc şi ni-L arată aşa cum era pe dinăuntru. Cei trei subliniază „omenescul\" din viaţa lui Isus, Ioan ne descopere „divinul\" din fiinţa şi lucrările Lui. Sinopticii sînt concentraţi asupra faptelor, Ioan ne lămureşte doctrinele.\r\n\r\nAceste caracteristici care ne coboară în intimitatea fiinţei Domnului Isus sînt factorii care ne fac să „simţim\" Evanghelia lui Ioan mai „diferit\" decît pe celelalte Evanghelii. Comentariile aşezate de Ioan în jurul „rostirilor\" Domnului Isus sînt adevărate studii care s-ar cere examinate cu cea mai mare atenţie, dar de care nu ne putem ocupa în spaţiul restrîns al acestei introduceri.\r\n\r\nCuvinte cheie şi teme caracteristice: O altă nestemată a Evangheliei lui Ioan este aşezarea întregului plan al mîntuirii într-o secvenţă de şapte miracole succesive săvîrşite de Domnul Isus. Evanghelistul ne avertizează că el nu a avut intenţia să scrie tot ceea ce ştie, ci că a „ales\" anumite întîmplări ca să ne ajute „să credeţi că Isus este Cristosul, Fiul lui Dumnezeu; şi crezînd, să aveţi viaţa în Numele Lui\"(Ioan 20:30-31)\r\n\r\n1.- Schimbarea apei în vin (Ioan 2)\r\n- „efectul primirii Cuvîntului într-o inimă de piatră\" (25)\r\n\r\n2.- Vindecarea fiului unui slujbaş (Ioan 4)\r\n- „acceptarea prin credinţă\" (4:50)\r\n\r\n3.- Vindecarea ologului (Ioan 5)\r\n- „un răspuns pozitiv la porunca divină\" (5:8)\r\n\r\n4. - Înmulţirea plinilor (Ioan 6)\r\n- „săvârşirea mîntuirii\" (6:32-33)\r\n\r\n5.- Umblarea pe mare (Ioan 6)\r\n- „un ajutor dincolo de puterile umane\" (6:21)\r\n\r\n6.- Vindecarea unui orb din naştere (Ioan 9)\r\n- „vedere nouă într-o lume nouă\" (9:3)\r\n\r\n7. - Învierea lui Lazăr (Ioan 11)\r\n- „nemuritor prin Isus Cristos\" (11:25)\r\n\r\nPrimele trei sînt ilustraţii pentru condiţiile mîntuirii, ultimele trei sînt consecinţele mîntuirii în viaţa celui credincios, iar cea de a patra este cheia întregii misiuni a lui Cristos. Pentru a înţelegere mai bună a Evangheliei lui Ioan, vă îndemnăm să citiţi şi comentariile făcute epistolelor scrise de el.\r\n\r\nSchiţa generală a Evangheliei lui Ioan urmăreşte tainic planul intrării Marelui Preot în CORTUL ÎNTILNIRII. Prin „iluminare\" divină Ioan înţelege că umbrele simbolice din Vechiul Testament au fost împlinite în viaţa Mîntuitorului. Cortul dat de Dumnezeu Israelului drept loc de întîlnire dintre divinitate şi oameni se transformă astfel într-o fascinantă anticipare a lucrării lui Dumnezeu, care, prin Cristos, a vrut să-şi „împace lumea cu Sine\" (2 Cor. 5:19).\r\n\r\nCei care au citit Evanghelia lui Ioan în limba greacă au rămas surprinşi şi încurcaţi de faptul că la începutul scrierii el proclamă:\r\n\r\n„Şi cuvîntul s-a făcut trup, şi a LOCUIT printre noi, plin de har şi de adevăr. Şi noi am privit slava Lui, o slavă întocmai ca slava singurului născut din Tatăl\" (Ioan 1:14)\r\n\r\nAcolo unde în limba română este „a locuit\", în limba greacă este „a cortuluit\", termen imposibil de tradus în limba noastră pentru că nu există. Ceea ce a vrut să facă Ioan a fost să atragă atenţia cititorilor că viaţa Domnului Isus va fi prezentată ca o „dezlegare\" a misterelor ascunse în tiparul Cortului Întîlnirii: viaţa lui terestră va fi prezentată drept lucrare de Mare Preot străbătînd mai întîi curtea de afară, apoi Sfînta cameră a Cortului şi pătrunzînd în final în Sfînta Sfintelor, cu propriul Său sînge pe care-l va turna pe Capacul Ispăşirii.\r\n\r\nÎn cîteva ocazii, Domnul Isus se identifică pe Sine cu acel „Eu sînt\", „Iehova\" prin care s-a făcut cunoscut Dumnezeu lui Moise. Cîteva pasaje în care Ioan afirmă categoric divinitatea Domnului Isus sunt: Ioan 1:1; 8:58; 10:30; 14:9, 20:28. Ca Dumnezeu întrupat, Isus Cristos se prezintă pe Sine prin cîteva expresii tipice: „Eu sînt pîinea vieţii\" (Ioan 6:35, 48), „Eu sînt lumina lumii\" (Ioan 8:12; 9:5), Eu sînt uşa (Ioan 10:7, 9), „Eu sînt păstorul cel bun\" (Ioan 10:11, 14), „Eu sînt învierea şi viaţa\" (Ioan 11:25), „Eu sînt calea, adevărul şi viaţa\" (Ioan 14:6), „Eu sînt adevărata viţă\" (Ioan 15:1-5).\r\n\r\nCUPRINSUL CĂRŢII\r\n\r\nI. Întruparea Fiului, 1:1-18\r\n\r\nII. Prezentarea Fiului, 1:19-4:54\r\nA. De către Ioan Botezătorul, 1:19-34\r\nB. pentru ucenicii lui Ioan, 1:35-51\r\nC. La nunta din Cana, 2:1-11\r\nD. La Templu, 2:12-25\r\nE. pentru Nicodim, 3:1-21\r\nF. Încă o dată de Ioan Botezătorul, 3:22-36\r\nG. pentru femeia samariteancă, 4:1-42\r\nH. pentru un slujbaş împărătesc, 4:43-54\r\n\r\nIII. Confruntările Fiului, 5-12\r\nA. La un praznic în Ierusalim, 5:1-47\r\n1. Semnul supranatural, 5:1-9\r\n2. Reacţia formalismului, 5:10-18\r\n3. Cuvîntarea, 5:19-47\r\n\r\nB. La un Paşte din Galileea, 6:1-71\r\n1. Semnul supranatural, 6:1-21\r\n2. Cuvîntarea, 6:22-40\r\n3. Reacţia celor fireşti, 6:41-71\r\n\r\nC. La un praznic al corturilor, 7:1-10:21\r\n1. Controversa nr.1: cuvîntarea, 7:1-29\r\n2. Reacţia ascultătorilor, 7:30-36\r\n3. Controversa nr.2: cuvîntarea, 7:37-39\r\n4. Reacţia ascultătorilor, 7:40-53\r\n5. Controversa nr.3: cuvîntarea, 8:1-58\r\n6. Reacţia ascultătorilor, 8:59\r\n7. Controversa nr.4:\r\nsemnul supranatural, 9:1-12\r\n8. Reacţiile, 9:13-41\r\n9. Controversa nr.5: cuvîntarea, 10:1-18\r\n10. Reacţii împărţite, 10:19-20\r\n\r\nD. La praznicul înnoirii Templului, 10:22-42\r\n1. Cuvîntarea, 10:22-30\r\n2. Refuzul şi acceptarea, 10:31-42\r\n\r\nE. La casa din Betania, 11:1-12:11\r\n1. Semnul supranatural, 11:1-44\r\n2. Reacţiile lor, 11:45-57\r\n3. Maria Îl unge cu mir, 12:1-8\r\n4. Reacţiile lor, 12:9-11\r\n\r\nF. La Ierusalim, 12:12-50\r\n1. Intrarea triumfală, 12:12-19\r\n2. Învăţăturile Lui, 12:20-50\r\n\r\nIV. În şcoala Fiului, 13:1-16:33\r\nA. Despre iertare, 13:1-20\r\nB. Despre trădare şi trădător, 13:21-30\r\nC. Despre plecarea Lui, 13:31-38\r\nD. Despre cer, 14:1-14\r\nE. Despre Duhul Sfînt, 14:15-26\r\nF. Despre pacea lăuntrică, 14:27-31\r\nG. Despre rodnicie, 15:1-17\r\nH. Despre lume ca sistem, 15:18-16:6\r\nI. Despre misiunea Duhului Sfînt, 16:7-15\r\nJ. Despre revenirea Lui, 16:16-33\r\n\r\nV. Mijlocirea Fiului, 17:1-26\r\n\r\nVI. Răstignirea Fiului, 18:1-19:42\r\nA. Arestarea, 18:1-11\r\n\r\nB.Judecata, 18:12-19:15\r\n1. Înaintea lui Ana, 18:12-23\r\n2. Înaintea lui Caiafa, 18:24-27\r\n3. Înaintea lui Pilat, 18:28-19:16\r\nC. Răstignirea, 19:17-37\r\nD. Înmormîntarea, 19:38-42\r\n\r\nVII. Învierea Fiului, 20:1-21:25\r\nA. Mormîntul gol, 20:1-10\r\n\r\nB. Arătările Domnului cel înviat, 20:11-21:25\r\n1. pentru Maria Magdalena, 20:11-18\r\n2. pentru ucenicii fără Toma, 20:19-25\r\n3. pentru ucenicii cu Toma, 20:26-31\r\n4. pentru şapte dintre ucenici, 21:1-14\r\n5. pentru Petru şi pentru ucenicul pe care-i iubea Isus, 21:25\r\n\r\nVIII. Infinitatea Fiului, 21:25" },
    //FAPTELE APOSTOLILOR
    { "FAPTELE APOSTOLILOR", "Explicatia Cartii: Fapte\r\n\r\nTitlul: În manuscrisele originale din limba greacă, cartea este numită: „Proxeis\" - „Faptele\". Scopul cărţii este arătat în chiar titlul ei: „Faptele Apostolilor\". Unii au fost de părere că titlul ar fi mai corect în alte variante ca: „Faptele Domnului Isus după înviere\" (ţinînd cont de cuvintele scrise de Luca în introducere: „...în cea dintîi carte am vorbit despre tot ce a început să facă Isus... pînă în ziua care s-a înălţat la cer\", concluzia ar fi că cea de a doua carte cuprinde faptele făcute de acelaş Isus după înălţarea la cer!). Alţii ar fi preferat: „Faptele Duhului Sfînt\" şi nu puţini sînt cei care remarcă, asemenea lui Marshall, că: „Nu le putem numi Faptele Apostolilor, atîta vreme cît, ignorînd activităţile celorlalţi apostoli, ele nu vorbesc decît despre unele fapte petrecute numai în vieţile lui Petru şi Pavel. Mai corect ar fi să le numim: „Unele fapte ale unor apostoli\"!\r\n\r\nPărerea noastră este că numele pe care-l poartă cartea este cel mai potrivit, atîta vreme cît asistăm la activităţile unui grup de oameni, trimişi de Isus Cristos în lume, îmbrăcaţi în puterea Duhului Sfînt şi care răspîndesc mesajul lor pe temelia mărturiei lor colective.\r\n\r\nAutorul: Este clar că autorul acestei cronici istorice este Luca, doctorul care a scris şi Evanghelia care-i poartă numele. Introducerea cărţii ne spune clar acest lucru:\r\n\r\n\"Teofile, în cea dintîi carte a mea, am vorbit despre ce a început Isus să facă şi să înveţe pe oameni, de la început, pînă în ziua în care S-a înălţat la cer, după ce, prin Duhul Sfînt, dăduse poruncile Sale apostolilor pe care-i alesese\" (Fapte 1:1-2).\r\n\r\nData: Luca a scris această cronică prin anii 62-63 d.Cr., pe la sfîrşitul celei dintîi detenţii a lui Pavel în Roma. Trebuie spus că Luca a scris probabil sub directa supraveghere a lui Pavel. Cartea nu poate fi plasată mai tîrziu, deoarece nu descrie nici persecuţia creştinilor din Roma sub Nero (64 d.Cr.), nici moartea marelui apostol (68 d.Cr.), şi nici căderea Ierusalimului (70 d.Cr.).\r\n\r\nConţinutul cărţii: Cartea Faptele Apostolilor este o trecere firească de la evenimentele Evangheliilor la conţinutul epistolelor. Din structura cărţii înţelegem că Duhul Sfint a inspirat scrierea acestei cronici din cel puţin trei motive majore:\r\n\r\n1. Să ilustreze programul de răspîndire al Evangheliei.\r\n\r\n„Şi-Mi veţi fi martori în Ierusalim, în toată Iudeea, în Samaria, şi pînă la marginile pămîntului\" (Fapte 1:8). Faptele Apostolilor debutează cu naşterea Bisericii în ziua de Rusalii şi se încheie odată cu ajungerea lui Pavel la Roma.\r\n\r\n2. Să adeverească autoritatea apostolică a lui Pavel.\r\n\r\nAceastă autoritate a fost mult comentată şi contestată în rîndul iudeilor. Luca scrie cronica cu gîndul de a arăta tuturor că Domnul Isus este Acela care l-a chemat pe Pavel, Acela care l-a învăţat şi Acela care l-a folosit pentru trimiterea Evangheliei la neamuri. Mai mult, Luca vrea să spună tuturor că Pavel este egalul lui Petru în lucrarea Evangheliei. Nu încape nici o îndoială că Luca simţea amărăciunea din sufletul lui Pavel. Epistolele acestui apostol al neamurilor sînt pline de pasaje în care el simte nevoia să se apere şi să se justifice. În multe locuri, persoana şi Evanghelia lui fuseseră vorbite de rău. Înarmat cu o mare doză de simpatie pentru Pavel şi martor la lucrările minunate pe care Dumnezeu le făcea prin apostol, Luca s-a aşternut la treabă hotărît să dovedească Bisericii că Dumnezeu i-a dat lui Pavel autoritate apostolică.\r\n\r\nPentru a-şi împlini acest de al doilea scop, Luca îşi aşează materialul scrierii într-un plan alcătuit minuţios şi echilibrat în desfăşurare. El alege din activităţile lui Petru şi Pavel exact acele evenimente care îi pun pe picior de egalitate. S-ar părea că Luca ar vrea să spună tuturor: „Voi îl respectaţi pe Petru? Ei bine, eu vă voi arăta că Pavel nu este cu nimic mai prejos\".\r\n\r\nPETRU PAVEL\r\nPrima predica (cap.2) Prima predica (13)\r\nVindecarea slabanogului (3) Vindecarea slabanogului (14)\r\nSimon Magul (8) Elima vrajitorul (13)\r\nPuterea umbrei lui (5) Basmale cu puteri vindecatoare (19)\r\nPunerea miinilor lui (8) Punerea miinilor lui (19)\r\nInchinarea catre Petru (10) Inchinarea catre Pavel (14)\r\nInvierea Tabitei (9) Invierea lui Eutih (2)\r\nIntemnitarea lui Petru (12) Intemnitarea lui Pavel (28)\r\n\r\n\r\nÎn situaţia în care Pavel a devenit autorul celor mai multe din epistolele Noului Testament, nu-i de mirare că s-a părut potrivit Duhului Sfînt să-l inspire pe Luca în scrierea Faptelor Apostolilor ca să ne arate tuturor cine este acest apostol şi cum a apărut el în sînul Bisericii primare. Fără scrierea lui Luca, epistolele lui Pavel ar fi fost mult mai greu de acceptat şi de asimilat.\r\n\r\n3. Pentru a arăta cea de a doua refuzare a împărăţiei de către evrei.\r\n\r\nDin predicile rostite de Petru la Rusalii şi imediat după aceea reiese foarte clar că dorinţa lui Dumnezeu ar fi fost ca evreii să se pocăiască şi să înceapă evenimentele prevestite de proorocul Ioel: „Şi acum ştim, fraţilor că din neştiinţă aţi făcut aşa, ca şi mai marii voştri. Dar Dumnezeu a împlinit astfel ce vestise mai dinainte prin gura tuturor proorocilor Lui: că adică, Cristosul Său va pătimi. Pocăiţi-vă dar, şi întoarceţi-vă la Dumnezeu, pentru ca să vi se şteargă păcatele, ca să vină de la Domnul vremuri de înviorare şi să trimeată pe Cel ce a fost rînduit mai dinainte pentru voi: pe Isus Cristos\" (Fapte 3:16-20). După învierea Sa din morţi, Domnul Isus a stat de vorbă cu ucenicii timp de 40 de zile despre: „lucrurile privitoare la Împărăţia lui Dumnezeu\" (Fapte 1:3). Rusaliile au fost începutul semnelor prevestitoare instaurării acestei împărăţii. Semnele, vindecările şi minunile au făcut parte şi ele din acest debut timpuriu al împărăţiei. Autorul Faptelor Apostolilor ne arătă însă că evreii nu au primit mesajul apostolilor şi cum, încet - încet, Dumnezeu s-a îndepărtat de la Israel, a făcut să înceteze iarăşi semnele împărăţiei şi l-a trimis pe Pavel „la neamuri\" (Fapte 22:21).\r\n\r\nCronica lui Luca acopere o perioadă de aproximativ 30 de ani din viaţa Bisericii primare. Ea consemnează trecerea de la Iudaism la credinţa personală în Isus Cristos ca Mîntuitor. Considerat la început doar ca o „sectă\" iudaică\", creştinismul se desprinde încet - încet de graniţele Israelului ducînd vestea despre Domnul Isus Cristos înspre marginile pămîntului.\r\n\r\nEvenimentele petrecute în primii 30 de ani din viaţa Bisericii sînt foarte importante. Doctrinele şi practica Bisericii primare s-au cristalizat în anii aceştia de înaintare a Evangheliei. Astăzi, pentru ca o învăţătură sau practică să poată fi considerată validă în Biserică, ea trebuie să fi fost subiectul unei învăţături a Domnului Isus şi o practică a Bisericii primare. Iată de ce cronica lui Luca a devenit filtrul care reglementează învăţătura şi practica bisericii creştine din toate timpurile.\r\n\r\nCuvinte cheie şi teme caracteristice: Cartea „Faptelor\" este o carte de tranziţie: de la Evanghelii la Epistole, de la Iudaism la Creştinism, de la Lege la Har, de la evrei în exclusivitate la orice făptură şi de la împărăţia lui Dumnezeu la Biserica lui Cristos.\r\n\r\nCUPRINSUL CĂRŢII\r\n\r\nI. Creştinismul în Ierusalim, 1-8:3\r\n\r\nA. Domnul cel înviat, 1:1-26\r\nDomnul Isus lămureşte, 1:1-5\r\nDomnul Isus trimite, 1:6-11\r\nDomnul Isus alege, 1:12-26\r\n\r\nB. Rusaliile: Naşterea Bisericii, 2:1-47\r\nPuterea de la Rusalii, 2:1-13\r\nPredica de la Rusalii, 2:14-36\r\nProzeliţii de la Rusalii, 2:37-47\r\n\r\nC. Vindecarea slăbănogului, 3:1-26\r\nMinunea, 3:1-11\r\nMesajul, 3:12-26\r\n\r\nD. Începutul prigoanei, 4:1-37\r\nPrigoana, 4:1-22\r\nRugăciunea, 4:23-31\r\nResursele, 4:32-37\r\n\r\nE. Compromis şi confruntare, 5:1-42\r\nRefuzarea compromisului intern, 5:1-11\r\nBiruinţa în confruntare externă, 5:12-42\r\n\r\nF. Noi lucrători pentru noi lucrări, 6:1-7\r\n\r\nG. Ştefan, cel dintîi martir, 6:8-8:3\r\nÎntărîtarea norodului, 6:8-15\r\nÎnfruntarea doctrinară, 7:1-53\r\nOmorîrea lui Ştefan, 7:54-8:3\r\n\r\nII. Creştinismul în Palestina şi Siria, 8:4-12\r\n\r\nA. Împrăştierea creştinilor, 8:4-40\r\nSamaritenii primesc Cuvîntul, 8:4-25\r\nEtiopeanul primeşte Cuvîntul, 8:26-40\r\n\r\nB. Convenirea lui Pavel, 9:1-31\r\nPavel îl primeşte pe Domnul, 9:1-19\r\nBiserica îl primeşte pe Pavel, 9:20-31\r\n\r\nC. Convertirea celor dintre neamuri, 9:32-11:30\r\nPregătirea lui Petru, 9:32-10:22\r\nPredica lui Petru, 10:23-48\r\nApărarea lui Petru, 11:1-18\r\nO nouă biserică la Antiohia, 11:19-30\r\n\r\nD. Irod îi prigoneşte pe creştini, 12:1-25\r\nOmorîrea lui Iacov, 12:1-2\r\nÎnchiderea şi eliberarea lui Petru, 12:3-19\r\nMoartea lui Irod, 12:20-23\r\nRăspîndirea Cuvîntului, 12:24-25\r\n\r\nIII. Creştinismul - către marginile pămîntului, 13:1-28:31\r\n\r\nA. Prima călătorie misionară, 13:1-14:28\r\nEvenimentele din Antiohia, 13:1-3\r\nEvenimentele din Cipru, 13:4-12\r\nEvenimentele din cetăţile Galatiei, 13:13-14:20\r\nEvenimentele din drumul de întoarcere, 14:20-28\r\n\r\nB. Consiliul de la Ierusalim, 15:1-35\r\nDeosebirea de vederi, 15:1-5\r\nDiscuţia, 15:6-18\r\nDecizia, 15:19-29\r\nComunicarea hotărîrii în Antiohia, 15:30-35\r\n\r\nC. A doua călătorie misionară, 15:36-18:22\r\nFormarea echipei misionare, 15:36-40\r\nA doua vizită la Biserici, 15:41-16:5\r\nTrecerea în Europa, 16:6-10\r\nLucrarea în Filipi, 16:11-40\r\nLucrarea în Tesalonic, Berea şi Atena, 17:1-34\r\nLucrarea în Corint, 18:1-17\r\nÎncheierea călătoriei, 18:18-22\r\n\r\nD. A treia călătorie misionară, 18:23-21:26\r\nLa Efes: Puterea Cuvîntului, 18:23-19:41\r\nÎn Grecia, 20:1-5\r\nAsia Mică: Troa şi presbiterii din Efes, 20:6-38\r\nDin Milet la Cezareea, 21:1-14\r\nPavel în Biserica din Ierusalim, 21:15-26\r\n\r\nE. Călătoria spre Roma, 21:27-28:31\r\nArestarea şi cuvîntarea de apărare, 21:27-22:29\r\nPavel înaintea Sinedriului, 22:30-23:10\r\nPavel escortat spre Cezareea, 23:11-35\r\nPavel se apără înaintea lui Felix, 24:1-27\r\nPavel se apără înaintea lui Festus, 25:1-27\r\nPavel se apară înaintea lui Agripa, 26:1-32\r\nCălătoria pe mare şi naufragiul, 27:1-44\r\nPavel în Malta şi spre Roma, 28:1-16\r\nPavel la Roma, 28:17-31" },
    //ROMANI
    { "ROMANI", "Explicatia Cartii: Romani\r\n\r\nEpistola către Romani este „magnum opus\"-ul apostolului Pavel. Niciuna dintre celelalte scrieri ale sale nu este la fel de bogată sau de organizată. În această scriere întîlnim deopotrivă „cunoştinţele\" lui Pavel şi „cunoştinţa\" lui în arta de a le aranja într-un tratat de teologie sistematică. Dintre toate cărţile scrise cîndva, aceasta este cea care a influenţat cel mai mult evoluţia şi cursul gîndirii creştine.\r\n\r\nFiecare creştin ar trebui să studieze conţinutul acestei epistole. Ea este „alfabetul\" credinţei, dar nu se opreşte aici, ci, în profunzimi, se dovedeşte busola tuturor călătoriilor prin doctrinele Bibliei. A deveni familiar cu învăţăturile aceste scrieri este tot una cu a fi „înrădăcinat\" în credinţă şi a dobîndi certitudini divine în problemele cruciale ale vieţii, morţii şi eternităţii.\r\n\r\nTitlul: În originalul grec, cartea poartă titlul: „Pro Romaious\" - „către Romani\". Cetatea Romei a fost fondată în anul 735 î.Cr. şi ajunsese pe vremea lui Pavel să fie cea mai măreaţă capitală a lumii, cu o populaţie de peste un milion de oameni.\r\n\r\nData: Pavel scrie această epistolă către creştinii din Roma în anul 57 d.Cr., pe cînd se afla în oraşul Corint, în casa lui Gaiu (Rom. 16:23; l Cor. 1:14). Scrierea a ajuns la Roma prin intermediul lui Fivi, diaconiţă a bisericii din Chencrea, unul din cartierele portuare ale Corintului (Rom. 16:1, 2).\r\n\r\nContextul istoric: Privind la vremea în care a fost scrisă această carte, vedem că nevoia după un aşa tratat doctrinar era evidentă. Cînd Pavel s-a apucat de scris trecuseră deja aproximativ 25 de ani de propovăduire a Evangheliei de-a lungul şi de-a latul Imperiului Roman. Comunităţile creştine se răspîndiseră pretutindeni, apărînd în toate colţurile Imperiului. Era inevitabil ca această nouă învăţătură să ridice întrebări chinuitoare în inimile celor ce o întîlneau pentru prima oară. Cum se împăca Evanghelia iertării cu „dreptatea\" lui Dumnezeu? Ce mai rămînea din „neprihănirea\" cerută de Dumnezeu dacă păcătoşii erau iertaţi prin oferta gratuită a „harului\"? Ce fel de relaţie era între această „Evanghelie\" şi străvechea „Lege a lui Moise\"? Nu-l desfinţa ea pe Moise? Şi ce mai rămînea din legămîntul „Avraamic\"? Cum se putea ca „neamurile\" să aibă parte de privilegiile aceluiaşi legămînt făcut cu evreii? Ce se va alege din nivelul moral, dacă se răspîndeşte acum vestea că Dumnezeu nu-i mai priveşte pe oameni prin filtrul pretenţiilor Legii, ci prin uşa deschisă a harului? Nu vor ajunge oare oamenii să creadă că este bine să păcătuim mai mult ca să se înmulţească şi mai mult harul? Ce mai rămînea valabil din promisiunile făcute de Dumnezeu Israelului? Mai rămînea în picioare statutul de popor al „legămîntului\"? Va mai avea Israelul, care-L respinsese pe Mesia, un rol în istoria viitoare a lumii? Nu cumva „legămîntul cel nou\" semnifica şi lepădarea Israelului ca popor? Multora dintre evreii evlavioşi li se părea că noua „Cale\" propusă de Pavel (Fapte 22:4; 24:22) aruncă pe fereastră tocmai tradiţiile şi moştenirile care le erau cele mai dragi.\r\n\r\nIată deci că pentru mulţi se cerea o explicaţie mai clară şi cumva definitivă asupra noii învăţături apărute în Biserica creştină.\r\n\r\nAutorul: Exista un singur om capabil de o asemenea lucrare! Dumnezeu pregătise deja un om care să primească această însărcinare! Lui Anania, Dumnezeu îi spusese:\r\n\r\n„Du-te, căci el (Pavel) este un vas, pe care l-am ales, ca să ducă Numele Meu înaintea neamurilor, înaintea împăraţilor, şi înaintea fiilor lui Israel.\" (Fapte 9:15)\r\n\r\nCu pregătirea lui temeinică de Fariseu, cu conştiinţa lui delicată şi cu credinţa lui puternică în religia evreilor, Pavel îşi dăduse seama, chiar mai mult decît împotrivitoni lui, de aspectul contradictoriu al învăţăturii ce-i fusese încredinţată. Dumnezeu a trebuit să-l convingă mai întîi pe el însuşi de temeinicia creştinismului. Faptul că a primit „Evanghelia lui\" direct prin revelaţie dumnezeiască (Gal. 1:1, 11-17), nu l-a scutit pe Pavel de zbucium lăuntric şi de multă frămîntare sufletească. În capitolul 9 el face mărturisirea: „simt o mare întristare şi am o durere necurmată în inimă... pentru fraţii mei, rudele mele trupeşti\" (9:2-3).\r\n\r\nS-ar cuveni să spunem acum ceva, pe scurt, despre apostolul Pavel. Dar cine poate spune ceva „pe scurt\" despre acest om extraordinar?\r\n\r\nIată ce scrie C.A. Fox despre marele apostol: „Mai multe calităţi, aparent contradictorii, au fost puse împreună de Dumnezeu pentru a împleti fiinţa lăuntrică a lui Pavel. Prin experienţa lui personală, el combină cunoştiinţe nemijlocite din cele trei sfere sociale care-i împărţeau pe oamenii din vremea lui. A fost ales din cea mai îngustă sectă a Iudaismului. Ca Fariseu, cunoştea legalismul evreu pe dinăuntru şi pe dinafară. A fost scos dintr-un mediu îmbibat cu cea mai aleasă cunoştinţă a culturii greceşti, căci şi-a trăit anii formării lui într-unul dintre cele mai importante centre de educaţie helenistică, Tarsul Ciliciei, şi şi-a însuşit temeinic eleganţa literaturii greceşti. Mai mult, încă de la naştere, s-a bucurat de privilegiile multiple ale celui cu „cetăţenia Romană\".\r\n\r\nPutem spune deci că Pavel a fost evreu pînă la măduvă, Grec în cel mai deplin sens al cuvîntului şi cetăţean Roman prin naştere. Dincolo de toate acestea, el a unit în personalitatea sa o neobişnuită vigoare intelectuală, o mare putere a voinţei, o simţire adîncă şi o mare compasiune pentru oameni.\r\n\r\nSpunînd toate acestea, mai trebuie să adăugăm un lucru, probabil cel mai important dintre toate. Iudaismul lui Pavel s-a frînt în întîlnirea directă cu Cristos pe drumul Damascului. Convertirea lui brusca şi capitularea lui necondiţionată, l-au făcut cel mai potrivit vas pentru a demonstra evreilor că Isus este viu şi că El este Mesia, Cel care trebuia să vină. Experienţa lui l-a ajutat să prezinte creştinismul nu ca pe ceva antagonist Iudaismului, ci ca pe o urmare firească, ca pe o continuare şi ca pe o împlinire a lui.\r\n\r\nDestinatarii epistolei: Cînd a scris această epistolă, Pavel încă nu fusese la Roma. Avea însă dorinţa aceasta arzătoare şi plănuia să călătorească într-acolo. Apostolul îşi dăduse seama de importanţa strategică a Bisericii stabilite în chiar capitala Imperiului. De acolo se puteau răspîndi apoi în toată lumea învăţăturile noii Evanghelii.\r\n\r\nFormată din evrei şi din neamuri, comunitatea creştină din Roma crescuse repede, probabil şi prin venirea multor creştini convertiţi din alte părţi ale Imperiului. Pavel îşi numeşte cititorii cînd evrei (2:17-29; 4:1; 7:1), cînd neamuri (1:13; 11:13-32; 15:15, 16, etc.). În încheierea scrisorii, el salută 26 persoane, dintre care două treimi au nume greceşti.\r\n\r\nConţinutul cărţii: Avînd în vedere componenţa Bisericii din Roma, Pavel construieşte o prezentare măiastră a adevărurilor măreţe ale creştinismului, care să-i dumirească şi pe evrei şi pe cei dintre neamuri. O schiţă a întregii scrisori ar cuprinde trei secţiuni terminate fiecare cu cîte o doxologie (8:38, 39; 11:33-36; 16:27-27)\r\n\r\nNu este nici o îndoială că primele 8 capitole sînt doctrinare ocupîndu-se cu doctrinele fundamentale ale Evangheliei. Secţiunea de la mijloc are un caracter naţional, clarificînd relaţia Israelului, ca naţiune, cu noua Evanghelie, iar ultima parte este o porţiune devoţională, care ilustrează pe scurt aplicarea noii învăţături la viaţa practică de toate zilele.\r\n\r\n1. DOCTRINAL: expoziţie - Cum mîntuieşte Evanghelia pe păcătoşi.\r\n2. NAŢIONAL: explicaţie - Cum se aplică Evanghelia Israelului.\r\n3. DEVOŢIONAL: aplicaţie - Cum se trăieşte Evanghelia.\r\n\r\nCuvinte cheie şi teme caracteristice: În faimoasa lui „Prefaţa (a studiul epistolei către Romani\", Martin Luther, celebrul reformator german, scrie următoarele:\r\n\r\n„Ca să pornim la drum, trebuie mai întîi să lămurim cîteva probleme de limbaj. Este absolut esenţial să nu începem studierea acestei epistole mai înainte de a vedea ce înţelege Sfîntul apostol Pavel prin termeni ca: Lege, păcat, har, credinţă, neprihănire, fire pămîntească, duh, etc. Fără lămurirea acestor termeni, citirea epistolei poate rămîne fără nici o valoare practică.\r\n\r\nIată de exemplu acest cuvînt mic: „Lege\". El nu trebuie luat în înţelesul lui social obişnuit care defineşte normele după care cetăţenii ştiu ce trebuie să facă şi ce trebuie să nu facă. Acest aspect este valabil numai în ce priveşte legile sociale omeneşti în care sînt judecate şi apreciate numai faptele, fără să se ţină în socoteală atitudinea inimii.\r\n\r\nLegea lui Dumnezeu este altfel. Dumnezeu judecă după ceea ce este în străfundurile inimii şi de aceea Legea Lui nu se opreşte la aspectul exterior al faptelor. Ea se pogoară în adîncimile fiinţei umane pretinzîndu-i nu numai un anumit fel de comportament, ci şi un anumit fel de simţire. Legea lui Dumnezeu pedepseşte chiar şi fapte aparent bune, atunci cînd acestea nu izvorăsc dintr-o pornire sinceră a inimii. Ipocrizia şi minciuna nu sînt tolerate.\r\n\r\nPsalmul 116:11 decretează: „Orice om este înşelător\", făcîndu-i pe toţi oamenii mincinoşi în comportamentul lor. Într-adevăr, omul nu poate ţine din toată inima Legea lui Dumnezeu. Este în natura noastră să ne placă parcă dinadins răul şi să nu ne atragă ceea ce este bine. Şi dacă inima noastră nu-şi găseşte plăcerea în ce este bine, atunci este clar că nici un om nu poate ţine în mod absolut Legea divină. Aceasta nu înseamnă altceva decît că omul este păcătos în străfundul naturii sale şi că această stare de păcat atrage asupra noastră mînia lui Dumnezeu, indiferent dacă aparent trăim ca nişte oameni respectabili făcînd în exterior fapte considerate de toţi ca fiind „bune\".\r\n\r\nConcluzia pe care o trage Sfîntul apostol Pavel în cuprinsul capitolului 2 din epistola către Romani este că pînă şi iudeii sînt păcătoşi, căci „nu cei ce aud Legea sînt neprihăniţi înaintea lui Dumnezeu, ci cei ce împlinesc legea aceasta, vor fi socotiţi neprihăniţi\" (2:13) El înţelege prin aceasta că niciunul, prin faptele făcute, nu este un împlinitor al legii. Dimpotrivă, el îi acuză pe toţi, direct în faţă, de comiterea celor mai flagrante încălcării ale Legii:\r\n\r\n„Tu care zici: „Să nu preacurveşti\", preacurveşti?\" (2:22).\r\n\r\n„Căci prin faptul că judeci pe altul, te osîndeşti singur; fiindcă tu, care judeci pe\r\naltul, faci aceleaşi lucruri\" (21).\r\n\r\nCe vrea să spună Pavel este aceasta: „Da, ştiu că tu trăieşti în exterior o viaţă care pare să respecte prevederile Legii şi că îi judeci pe ceilalţi care nu fac la fel ca tine; ştiu că eşti foarte priceput să vezi paiul din ochiul aproapelui tău, dar de bîrna care îţi împiedică vederea n-ai habar!\"\r\n\r\n„Adevărul este că şi dacă ţii în aparenţă Legea, cu fapte exterioare, din pricina pedepsei sau de dragul răsplătirii, în lăuntrul fiinţei tale faci lucrarea aceasta fără nici o plăcere, împotriva pornirilor naturale şi numai împins de la spate. Dacă ţi s-ar da voie, ai face deîndată exact ceea ce acum Legea te opreşte.\"\r\n\r\n„Concluzia firească care se impune este că în lăuntrul tău, tu urăşti Legea. Atunci ce importanţă mai are că tu îi înveţi pe alţii să nu fure, cînd ştim că tu ai hoţia în inimă, şi ai fura din toată inima dacă te-ai lăsa dus de pornirile inimii tale? Nu vezi cîtă ipocrizie se ascunde în dosul unei măşti de decenţă? Tu îi înveţi pe alţii, dar nu te poţi convinge nici pe tine însuţi. De fapt, respectînd ceea ce tu, ca evreu, numeşti Lege, n-ai ajuns s-o şi înţelegi vreodată.\"\r\n\r\nLegea nu ne poate rezolva problemele, în capitolul 5 al epistolei, Sfîntul Pavel spune clar că Legea nu a venit ca să ne facă mai buni, ci doar ca să scoată în evidenţă păcatul. Cu cît Legea ne pretinde mai multe, cu atîta noi ne împotrivim ei, urînd-o din toată inima.\r\n\r\nAcesta este motivul pentru care, Pavel spune în capitolul 7:14 că: „Legea este duhovnicească, dar eu sînt... rob păcatului\". Cum se înţelege aceasta? Iată cum: dacă Legea ar fi fost dată numai pentru trupul exterior, ea ar fi putut fi satisfăcută prin fapte exterioare; dar atîta vreme cît Legea este duhovnicească, nimeni nu o poate împlini, pentru că ne este împotrivă firea noastră păcătoasă. Numai Dumnezeu ne poate schimba inima şi ne poate dărui una care îl poate ridica pe om la nivelul Legii lui Dumnezeu. Numai prin lucrarea de înnoire făcută de Duhul Sfînt putem ajunge să dorim să facem voia lui Dumnezeu, nu de frica pedepsei sau din obligaţie, ci dintr-o pornire sinceră a inimii.\r\n\r\n„Legea, care este duhovnicească\" nu poate fi împlinită decît de un om făcut „duhovnicesc\" prin lucrarea de înnoire a Duhului. Acolo unde Duhul Sfînt încă nu a intrat, rămîne în continuare păcatul, rămîne împotrivirea tacită faţă de Lege şi duşmănia faţă de prevederile ei. Aceasta se întîmplă cu toate că, mintal, noi recunoaştem că voia lui Dumnezeu este bună, dreaptă şi sfîntă.\r\n\r\nCăutaţi să vă familiarizaţi cu felul acesta de gîndire al lui Pavel şi veţi ajunge să vă daţi singuri seama că „făcînd faptele Legii\" şi „împlinirea Legii\" sînt două lucruri cît se poate de deosebite. Faptele Legii sînt însumarea a tot ceea ce face cineva din dorinţa sinceră de a respecta perceptele divine prin strădaniile puterilor proprii. Oricît de sincere sînt însă aceste strădanii, ele sînt însoţite mereu de o stare de înverşunare a inimii şi de o constrîngere a pornirilor lăuntrice, care fac în ultimă instanţă faptele exterioare tot una cu ipocrizia, golindu-le de orice valoare. Aceasta este cauza care-l face pe Pavel să scrie în Romani 3:20:\r\n\r\n„Căci nimeni nu va fi socotit neprihănit înaintea Lui, prin faptele Legii, deoarece prin Lege vine cunoştinţa deplină a păcatului\".\r\n\r\nCît de caraghioşi sînt unii care-i învaţă pe cei din Biserici „să se pregătească să primească harul făcînd faptele necesare\"! Cum ar putea un om să se pregătească făcînd faptele, atîta timp cît inima i se împotriveşte şi-l umple de fiere amară? Cum s-ar putea ca o astfel de „faptă\" făcută din obligaţie sau constrîngere a inimii să fie plăcută înaintea lui Dumnezeu?\r\n\r\nPe de altă parte, a împlini Legea înseamnă a face faptele ei din dragoste, cu o inimă voioasă şi binevoitoare, fără a simţi presiunea necesităţii sau apăsarea constrîngerii. Această atitudine voioasă şi binevoitoare a inimii este produsul lucrării Duhului Sfînt în lăuntrul celui mîntuit:\r\n\r\n„...pentru că dragostea lui Dumnezeu a fost turnată în inimile noastre prin Duhul Sfînt, care ne-a fost dat\" (5:5).\r\n\r\nDar Duhul Sfînt nu este dat decît „prin\", „în urma\" şi „ca rezultat\" al credinţei mîntuitoare în Domnul Isus Cristos. Aceasta este ceea ce spune Pavel în introducerea sa:\r\n\r\n„Pavel... pus deoparte să vestească Evanghelia lui Dumnezeu, pe care o făgăduise mai înainte prin proorocii Săi în Sfintele Scripturi.\r\n\r\n„Ea priveşte pe Fiul Său, născut din sămînţa lui David, în ce priveşte trupul, iar în ce priveşte duhul sfinţeniei dovedit cu putere că este Fiul lui Dumnezeu, prin învierea morţilor; adică pe Isus Cristos, Domnul nostru, prin care am primit harul...” (Romani 1:1-5).\r\n\r\nSchimbarea inimii se face prin lucrarea Duhului, care la rîndul Său nu este dat decît ca rezultat al credinţei. Putem spune aşadar fără nici o ezitare că singurele fapte bune sînt cele care sînt o consecinţă a credinţei. Numai „neprihănirea care se capătă prin credinţă\" poate împlini Legea, căci numai din meritul cîştigat de Cristos primim în dar lucrarea Duhului care ne transformă inima făcînd-o să-i placă lucrurile cerute de Dumnezeu.\r\n\r\nAceasta este de altfel şi definiţia harului, care nu este un fel de certificat pentru libertatea de a face tot ceea ce ne pofteşte inima păcătoasă. Harul este dorinţa şi puterea pe care ne-o dăruieşte Dumnezeu pentru a-l împlini voia Sa sfîntă.\r\n\r\nCredinţa nu desfinţează faptele Legii, ci le face accesibile celui mîntuit:\r\n\r\n„Deci, prin credinţă desfinţăm noi Legea? Nicidecum. Dimpotrivă, noi întărim Legea\" (Rom. 3:31).\r\n\r\n„Păcatul\" este un alt termen care trebuie explicat. În Sfînta Scriptură, păcatul nu este numai săvârşirea faptelor care încalcă prevederile Legii, ci denumeşte un întreg complex de sentimente şi atitudini ale inimii care ne îndemnă să călcăm Legea.\r\n\r\nDumnezeu nu se opreşte la aparenţe, ci pune degetul direct pe rană atunci cînd „nu se uită la ce izbeşte ochiul, ci priveşte la inimă\". Înainte de a deveni fapte, păcatele noastre erau ascunse în cutele infinit de sensibile ale inimii:\r\n\r\n„Dar ce iese din gură, vine din inimă, şi aceea spurcă pe om. Căci din inimă ies gîndurile rele, uciderile, preacurciile, curviile, furtişagurile, mărturiile mincinoase, hulele. Iată lucrurile care spurcă pe om\" (Matei 15:17-20).\r\n\r\nPrin urmare, credinţa este singura noastră cale spre neprihănire, căci credinţa ne aduce în inimă lucrarea Duhului Sfînt.\r\n\r\nÎn Ioan 16:8-9, Domnul Isus spune că singurul păcat care nu li se va ierta oamenilor este necredinţa:\r\n\r\n„Cînd va veni El (Duhul Sfînt) va dovedi lumea vinovată în ce priveşte păcatul, neprihănirea şi judecata.\r\n\r\nÎn ce priveşte păcatul: fiindcă ei nu cred în Mine\".\r\n\r\nCel ce crede are deschisă calea spre neprihănirea pornită dintr-o inimă spălată de sîngele Domnului Isus şi înnoită de lucrarea transformatoare a Duhului. Dintr-o astfel de inimă vor curge apoi „rîuri de apă vie, cum zice Scriptura\" (Ioan 7:38).\r\n\r\nÎnainte de a exista fapte bune sau rele, există o inimă stăpînită de credinţă sau de necredinţă Inima firii pămînteşti este rădăcina tuturor relelor. Ea este„capul şarpelui\" despre care vorbeşte Scriptura şi despre care i-a fost promis lui Adam că va fi zdrobit sub picioarele seminţei lui (Gen 3-15)\"\r\n\r\nCUPRINSUL CĂRŢII\r\n\r\nEPISTOLA CĂTRE ROMANI - „Neprihănirea dată-n dar\"\r\n\r\nINTRODUCERE (1:1-17)\r\n\r\n1. Neprihănirea lui Dumnezeu este ceea ce ne trebuie\r\nl:18-1:32 Omenirea L-a părăsit pe Dumnezeu şi a pierdut neprihănirea.\r\n2:1-2:16 neamurile, fără Lege, sînt vinovate.\r\n2:17-3:8 iudeii, sub Lege, sînt şi ei vinovaţi.\r\n3:9-3:20 Rezultatul: Vinovăţia întregii omeniri.\r\n\r\n2. Neprihănirea lui Dumnezeu ne este dăruită\r\n3:21 Nu prin ţinerea Legii.\r\n3:22, 23 Prin credinţă (calea).\r\n3:24 Din harul divin acordat (izvorul).\r\n3:24, 25 În urma morţii lui Cristos (prilejul).\r\n3:26-31 Fără plată pentru neamuri, ca şi pentru evrei (scopul).\r\n4:1-25 Exemplificată în Avraam şi în David (exemplele).\r\n5:1-11 Are ca urmare mari binecuvîntări (rezultatul).\r\n5:12-21 Rezumatul: Mîntuirea este oferită întregii omenirii.\r\n\r\n3. Neprihănirea lui Dumnezeu este realizată în noi\r\n6:1-13 Noi am murit faţă de păcat şi trăim pentru Dumnezeu.\r\n6:14-7:25 Noi am murit faţă de Lege şi trăim sub har.\r\n8:1-13 Noi am murit faţă de fire şi trăim prin Duhul.\r\n\r\n4. Neprihănirea lui Dumnezeu ne este asigurată\r\n8:14-25 Sîntem fii şi ne aşteaptă slava.\r\n8:26-27 Mijlocirea Duhului\r\n8:28-34 Scopul etern al Tatălui\r\n8:35-39 Dragostea statornică a Fiului.\r\n\r\n5. Neprihănirea lui Dumnezeu este la lucru pentru noi\r\nCap.9 Israelul a fost ales în trecut.\r\nCap. 10 Israelul este îndepărtat în prezent.\r\nCap.11 Israelul va fi restaurat în viitor.\r\n\r\n6. Neprihănirea lui Dumnezeu este arătată prin noi\r\nCap.12, 13 În cadrul activităţilor spirituale, sociale şi cetăţeneşti.\r\nCap.14-16 În cadrul părtăşiei şi slujirii noastre creştine." },
    //1 CORINTENI
    { "1 CORINTENI", "Explicatia Cartii: 1 Corinteni\r\n\r\nProbabil că nici una dintre epistolele apostolului Pavel nu a fost scrisă unui grup mai frămîntat de probleme, de compromisuri cu păcatul şi de lupte ca această scrisoare adresată creştinilor din oraşul Corint. Au existat glasuri care au spus că Biserica din Corint poate fi numai în parte şi cu greu considerată o Biserică „creştină\". Totuşi, faptul că Pavel o numeşte aşa şi mai ales faptul că Duhul Sfînt a socotit că scrisorile adresate de apostol credincioşilor de acolo merită să fie păstrate în canonul Noului Testament, ne îndeamnă să avem o altă părere. Este bine să stăruim cu atenţie asupra acestei Biserici şi să vedem ce mesaj găsim în epistolele adresate ei pentru viaţa Bisericilor de astăzi.\r\n\r\nTitlul: În originalul grec, cartea se numeşte: „Pros Korinthious A\" - „Către Corinteni A\" (sau „întîia\").\r\n\r\nAutorul: Pavel este nu numai autorul acestei epistole, ci şi fondatorul Bisericii din Corint 1corinteni. 4:14-15\r\n\r\nContextul scrierii: Corintul era un mare centru comercial, cultural, religios şi vai, un mare centru al desfrîului. Intrarea apostolului în oraş s-a petrecut la mai bine de o sută de ani după ce Iulius Cezar reconstruise cetatea, transformînd-o într-o nouă capitală a Ahaiei. Din punct de vedere maritim, Corintul era un oraş situat pe promontoriul dintre două porturi: Chencrea la est şi Laceum la vest. Această poziţie i-a asigurat accesul traficului maritim comercial din toată lumea. Cînd spuneai Corint, spuneai afluenţă materială, tranzacţii comerciale, garnizoane militare, lux, afluenţă materială şi... iar afluenţă materială. Cînd Pavel aminteşte de: „aur, argint şi pietre preţioase\" în capitolul 3, el vorbeşte cu oameni care cunoşteau foarte bine valoarea acestor mărfuri.\r\n\r\nLa 16 kilometri de porţile cetăţii se întindeau cîmpurile destinate Jocurilor Istmice, un corespondent al Jocurilor Olimpice de astăzi. Din patru în patru ani veneau acolo atleţi din toată lumea să se întreacă în tot felul de jocuri dintre care cursele, boxul şi luptele erau cele mai renumite. Pavel face aluzie la aceste întreceri în capitolul 9.\r\n\r\nSus pe înălţimea Acropolis, dominînd şi la propriu şi la figurat cetatea, se înălţa mîndru Templul Afroditei, zeiţa iubirii şi a fertilităţii. O mie de preotese practicau prostituţia ca parte a ritualului de închinăciune. Corintul era un oraş al viciului. Pe străzi se plimbau bărbaţi homosexuali care-şi lăsaseră părul să crească ca la femei. Vorbind despre ei, Pavel scrie în l Cor. 11:4:\r\n\r\n„Nu vă învaţă chiar şi firea că este ruşine pentru un bărbat să poarte părul lung...” Ne putem închipui ce mare trebuie să fi fost dezgustul lui Pavel la vederea decadenţei morale din jur. Nici în Antiohia nu întîlnise o aşa combinaţie de „înţelepciune\" lumească şi imoralitate animalică. În toată Biblia nu se găseşte o descriere mai vie a stării de păcat a omenirii decît aceea făcută de Pavel în capitolul 1 al epistolei către Romani. Este suficient să spunem că acel comentariu trist fusese scris pe vremea în care apostolul se afla în cetatea Corintului.\r\n\r\nPavel a venit la Corint după ce vizitase Atena. Acolo avusese o experienţă nu prea încurajatoare cu cei ce-l ascultaseră (Fapte 17:15-34). Probabil că starea lui sufletească nu era prea bună. Grecia nu părea să fie un cîmp bun pentru Evanghelie. După îngîmfata Atena, a urmat decăzutul Corint! Nu-i de mirare că Dumnezeu a trebuit să i se arate noaptea şi să-l încurajeze:\r\n\r\n„Noaptea, Domnul a zis lui Pavel într-o vedenie: „Nu te teme; ci vorbeşte şi nu tăcea, căci Eu sînt cu tine; şi nimeni nu va pune mîna pe tine, ca să-ţi facă rău: vorbeşte, fiindcă am mult norod în această cetate\". Aici a rămas un an şi şase luni, şi învăţa printre Corinteni Cuvîntul lui Dumnezeu\" (Fapte 18:9-10).\r\n\r\nPrimii cu care s-a întîlnit Pavel au fost Aquila şi Priscila, evrei creştini, victime a expulzării evreilor din Roma sub Claudiu Cezar. Pentru că se ocupau şi ei cu facerea corturilor, casa lor a devenit şi casa lui Pavel. După ce Sila şi Timotei, care fuseseră în Macedonia, au întregit echipa misionară, Pavel s-a dedat cu totul propovăduirii, fiind susţinut material de ceilalţi (Fapte 18:5).\r\n\r\nÎn Fiecare zi de Sabat el predica în Sinagogă „dovedind iudeilor că Isus este Cristosul\". Două convertiri notabile s-au produs ca urmare a predicării: Iust, un evreu a cărui casă era vecină cu Sinagoga şi-n casa căruia s-a mutat Pavel, şi Crisp, fruntaşul Sinagogii, care s-a întors la Domnul cu toată casa lui (Fapte 18:7-8) Aşa a luat fiinţă Biserica Nou Testamentală din Corint alcătuită din bărbaţi şi femei, evrei şi dintre neamuri, sclavi şi oameni liberi. Despre caracterul şi caracteristicile acestei adunări găsim ceea ce spune Pavel în cap.l:26-31 şi 6:9-11:\r\n\r\n\"De pildă, fraţilor, uitaţi-vă la voi care aţi fost chemaţi: printre voi nu sînt mulţi înţelepţi în felul lumii, nici mulţi puternici, nici mulţi de neam ales...”, „Nu ştiţi că cei nedrepţi nu vor moşteni împărăţia lui Dumnezeu? Nu vă înşelaţi în privinţa aceasta: nici curvarii, nici închinătorii la idoli, nici preacurvarii, nici malahii, nici sodomiţii, nici hoţii, nici cei lacomi, nici beţivii, nici defăimătorii, nici răpăreţii nu vor moşteni Împărăţia lui Dumnezeu. Şi aşa eraţi unii din voi! Dar aţi fost spălaţi, aţi fost sfinţiţi...”\r\n\r\nConţinutul cărţii: 1 Corinteni este o scrisoare plină de mînie, de mustrare, de corectare şi de învăţătură. După cele 18 luni petrecute în Corint, Pavel a plecat la Efes unde a stat 3 ani de zile. Fiind acolo, el a primit o scrisoare din partea credincioşilor din Corint în care i se cerea părerea despre căsătorie şi despre carnea rămasă de la jertfele păgîne. Trei foarte cunoscuţi membrii ai adunării din Corint i-au adus lui Pavel scrisoarea la Efes: Ştefanas, Fortunat şi Ahaic (1 Cor. 16:17, 18). Nu încape nici o îndoială că între Pavel şi aceşti trei fraţi din Corint au avut loc discuţii amănunţite din care apostolul a aflat despre starea decăzută a credincioşilor Corinteni. Răspunsul lui Pavel este îndreptat deci nu numai spre problemele ridicate de ei, ci înspre combaterea şi corectarea tuturor relelor despre care aflase.\r\n\r\nSituaţia ar putea fi descrisă în aceste cuvinte: Cei din Corint fuseseră cu adevărat întorşi la Cristos şi formau acum o adunare de copii ai lui Dumnezeu. Ei o rupseseră în teorie pentru totdeauna cu idolatria şi cu practicile idolatre. Totuşi, ei nu puteau, peste noapte, să se desprindă de ceea ce fusese aşezat an după an „în ei\" şi nu ştiau cum să se deslipească de ceea ce continua să se întîmple „în jurul lor\". Practica vieţii lor nu se ridicase la nivelul „teoriei\".\r\n\r\nMisionarii de astăzi ne povestesc despre situaţii asemănătoare în care se găsesc unii convertiţi din ţările păgîne. Acceptarea lui Cristos este amestecată adeseori cu forme tradiţionale de idolatrie străveche. Oamenii nu se pot desprinde imediat de „formele\" care le-au dat identitatea naţională timp de secole. Se ajunge astfel la o stare de impas, pe care unii încearcă să o depăşească pe calea compromisurilor. Aşa şi-au făcut loc în Biserică, de-a lungul veacurilor, tot felul de sărbători şi de obiceiuri păgîne ca: rugăciunea pentru cei morţi, teama de spiritele celor morţi, pomana pentru sufletul morţilor, cultul regenerării de la schimbarea anilor, descîntecele şi sărbătorirea unor zile preluate din calendarele păgîne. În relaţia cu păcatul, compromisul înseamnă însă „robie\" spirituală (6:12). Corintenii au avut nevoie ca cineva să le spună că nu există nici o cale de mijloc: ori cu Dumnezeu, ori cu lumea păgînă. O alegere trebuia făcută şi această alegere trebuia făcută repede. Acesta este în esenţă mesajul din l Corinteni 5 şi 6. Duhul Sfînt este întristat într-o adunare în care este tolerat păcatul:\r\n\r\n„Nu ştiţi că voi sînteţi Templul lui Dumnezeu şi că Duhul lui Dumnezeu locuieşte în voi?... căci Templul lui Dumnezeu este sfînt şi aşa sînteţi şi voi\"(3:16-17}...”să fi fost dat afară din mijlocul vostru\", „Daţi afară dar din mijlocul vostru pe răul acela\"(5:2, 13) „Căci aţi fost cumpăraţi cu un preţ. Proslăviţi dar pe Dumnezeu în trupul vostru şi în duhul vostru, care sînt ale lui Dumnezeu” (6:20)\r\n\r\nCuvinte cheie şi teme caracteristice: Una dintre cele mai frumoase lecţii pe care le dă Pavel Corintenilor este lecţia exemplului personal. Ca reprezentant al lui Cristos într-o lume pierdută, apostolul se dă pe sine pildă celor din greu încercata Biserică a Corintului. Sub presiunea lumii, credincioşii din toate timpurile au avut nevoie de lideri care nu numai să le spună ce au de făcut, ci să le şi arate cum trebuie trăită viaţa creştină. În capitolul 4:16 el spune: „de aceea vă rog să călcaţi pe urmele mele\", adăugînd în capitolul 11:1: „Călcaţi pe urmele mele, întrucît şi eu calc pe urmele lui Cristos\".\r\n\r\nIată cele zece pasaje în care Pavel se dă drept exemplu demn de urmat:\r\n\r\n1. Loialitate în mesaj, metodă şi motivaţie (2:1-5).\r\n2. Seriozitate în punerea temeliei şi în zidirea de deasupra (3:10-23).\r\n3. Credincioşie în lucrul încredinţat lui (4:1-6)\r\n4. Răbdare în suferinţele pentru Cristos. 4:9-16\r\n5. Consideraţie pentru fratele mai slab. 6:12; 8:13\r\n6. Renunţare la drepturi şi privilegii. 9:12-18\r\n7. Negare de sine pentru cîştigarea altora. 9:19-23\r\n8. Autodisciplinare a trupului şi a comportamentului. 9:27; 10:33\r\n9. Stăpînire de sine în adunările publice. 14:18-20\r\n10. Rîvnă şi recunoştinţă. 15:9-10\r\n\r\nCUPRINSUL CĂRŢII\r\nIntroducere (1:1-9)\r\n\r\nI. MUSTRARE PENTRU DEZBINARE 1 - 6\r\n(Corintenii se complăceau în a da slavă oamenilor -1:12)\r\n\r\nPartidele ataşate oamenilor sînt rele:\r\n- pentru că mîntuirea prin Cruce înlătură orice fel de înţelepciune omenească (v.18-31)\r\n- deoarece adevărata înţelepciune este dată de Duhul Sfînt; ea nu aparţine oamenilor (v.5-13)\r\n- deoarece „învăţătorii\" nu sînt decît „slujitori\"; puterea este a Domnului (3:5, 6, 21; 4:1)\r\n- atîta timp cît această „fală\" (5:2) este ipocrită (5:6), căci credincioşii continuau să se complacă în mijlocul unor păcate murdare ca incestul, dările în judecată şi abuzurile de tot felul.\r\n\r\nII. RĂSPUNSURI LA PROBLEME RIDICATE 7-15\r\n(Corintenii i-au scris lui Pavel despre aceste probleme - 7:1)\r\ncap. 7 - Răspuns la problema: „căsătorie sau celibat?\"\r\ncap. 8-10 - Despre carnea rămasă de la jerfele păgîne.\r\n\r\nPrincipiul (cap.8), exemplul lui Pavel (9),\r\navertismente din Scriptură (10),\r\nrezolvarea (10:23-11:1)\r\n\r\ncap.11 - Ţinuta femeii în adunare (v.2-16) şi comportarea la Cina Domnului (v.17-34).\r\ncap. 12-14 - Răspuns în problema darurilor spirituale. Distribuite de Duhul (12), sărace fără dragoste (13), inferioare profeţiei (14).\r\ncap. 15 - Răspuns la problema învierii sfinţilor. Relaţia cu învierea lui Cristos (v. 1 -19), perspectiva (v.20-34), trupurile celor înviaţi (v.35-49), „taina\" (v.50-58).\r\nApendice - cap. 16" },
    // 2 CORINTENI
    { "2 CORINTENI", "Explicatia Cartii: 2 Corinteni\r\n\r\n„V-am scris cu multă mîhnire, şi cu strîngere de inimă, cu ochii scăldaţi în lacrimi, nu ca să vă întristaţi, ci ca să vedeţi dragostea nespus de mare pe care o am faţă de voi\" (2 Cor. 2:1).\r\n\r\nTitlul: în originalul grec, cartea se numeşte: „Pros Korinthious B\"- „Către Corinteni B \" (sau a doua\").\r\n\r\nAutorul: S-ar putea ca alte epistole să fie mai profunde decît aceasta, dar niciuna dintre ele nu este mai caldă şi mai personală ca ea. În niciuna din celelalte nu se revarsă inima lui Pavel mai plenar ca în aceasta şi din nici o alta nu aflăm mai multe despre viaţa şi frămîntările apostolului.\r\n\r\nData: Asa cum am văzut deja, cea dintîi epistolă a lui Pavel către Corinteni a fost scrisă pe cînd apostolul se află în Efes (1 Cor.l6:8). La scurt timp după scrierea ei, el a fost obligat să fugă din cetate din pricina răscoalei puse la cale de argintarul Dimitrie şi de ceata lui idolatră. Zelul pentru zeiţa lor Artemis (Diana) şi mai cu seamă preocuparea pentru cîştigul lor de făcători de machete ale Templului păgîn, i-a ridicat pe aceştia împotriva lui Pavel şi a Evangheliei (Fapte 19:21-41).\r\n\r\nDe la Efes, apostolul a plecat înspre Troa, apoi a trecut prin provinciile nord-estice ale Mării Egee pentru a ajunge în ţinuturile Macedoniei. După ce a vizitat bisericile de acolo, Pavel s-a îndreptat spre sud, ajungînd iarăşi în Grecia, vizitînd din nou Ahaia şi Corintul şi zăbovind acolo încă trei luni de zile (Fapte 20:1-3). Era a treia vizită pe care avea să le-o facă celor din Corint (2 Cor. 12:14; 13:1).\r\n\r\nÎn intervalul de timp scurs între plecarea şi reîntoarcerea în Corint, Pavel a scris această a doua epistolă. Locul scrierii a fost probabil Filipi, iar circumstanţele prin care trecea apostolul au fost dintre cele mai chinuitoare.\r\n\r\nContextul scrierii: Pavel scrie aceste rînduri dintr-una din cele mai grele şi mai întunecate situaţii în care s-a aflat vreodată. Apostolului i-a fost dat să experimenteze necazul şi deznădejdea pentru ca să ne poată scrie nouă şi să ne ajute să trecem la rîndul nostru prin clipele noastre de încercare (2 Cor. 1:3-4).\r\n\r\nLui Pavel nu-i era indiferent ceea ce se putea petrece în Corint. Credinciosul său tovarăş de lucru, Tit, ar fi trebuit să-l aştepte deja la Filipi, cu un raport despre cele petrecute în cetate după plecarea sa. Faptul că el întîrzia să apară nu făcea decît să sporească şi mai mult neliniştea şi temerile apostolului (2 Cor. 2:12, 13). De fapt, asupra lui Pavel se aruncaseră parcă dintr-o dată toate îndoielile, toate dezamăgirile şi toate atacurile diavolului, cufundîndu-l pe încercatul apostol într-unul din cele mai negre şi mai sfîşietoare momente din întreaga lui carieră de luptător pentru propăşirea şi apărarea Evangheliei:\r\n\r\n„Căci şi după venirea noastră în Macedonia, trupul nostru n-a avut nici o odihnă. Am fost necăjiţi în toate chipurile: de afară lupte, dinăuntru temeri\" (7:5)\r\n\r\nSe părea că cei din Corint erau cu totul porniţi împotriva lui. Pe de altă parte Galatia căzuse pradă ereziilor iudaizatoare şi acceptaseră „o altă Evanghelie\" (Galat. 1:6-9). În Efes, scăpase ca prin urechile acului din mîinile lui Dimitrie şi a celor ce au aţîţat cetatea împotriva lui. Se luptase cu oameni, dar lupta fusese atît de cruntă încît apostolul o aseamănă cu o luptă de gladiatori împotriva, „fiarelor\" (1 Cor. 15:32). Gîndul că îşi lăsase mica turmă de urmaşi ca pe nişte miei în mijlocul lupilor nu-i da odihnă. Sub acest potop de grijuri, sănătatea lui a cedat din nou făcînd loc unui atac de boală, care ameninţa să-i curme viaţa:\r\n\r\n„În adevăr, fraţilor, nu voim să vă lăsăm în necunoştinţă despre necazul care ne-a lovit în Asia, de care am fost apăsaţi peste măsură de mult, mai pe sus de puterile noastre, aşa că nici nu mai trăgeam nădejde de viaţă. Ba încă ne spunea gîndul că trebuie să murim...” (2 Cor. l:8-9).\r\n\r\nDumnezeu avea însă alte planuri cu Pavel. Tit soseşte din Corint cu veşti îmbucurătoare. Ca un răspuns dat primei epistole a lui Pavel, prin biserica din Corint trecuse focul curăţitor al pocăinţei. Asprimea condamnărilor apostolului, disciplina anunţată de el, mustrarea şi atenţionarea căzuseră pe un pămînt bun, pregătit de Dumnezeu şi care nu zăbovise să dea la iveală roade de pocăinţă (2 Cor. 7:6:16).\r\n\r\nÎntr-atît de bucuros a fost Pavel de raportul dat de Tit, încît s-a aşternut imediat la scris şi le-a trimis această a doua scrisoare. Tit a călătorit cu ea înapoi la Corint ca să se desăvârşească lucrarea bună începută acolo (8:16, 17, 23).\r\n\r\nCuprinsul cărţii: Oricît de îmbucurător a fost raportul dat de Tit lui Pavel, el a cuprins şi aspecte problematice care nu puteau fi trecute cu vederea. Chiar dacă cei din Corint dăduseră dovadă că au o mare rîvnă în ce priveşte sfinţenia, nu se poate spune că toţi au primit sfaturile lui Pavel. Pavel îşi exprimă poziţia faţă de aceştia în mod clar:\r\n\r\n„De multă vreme voi vă închipuiţi că vrem să ne apărăm înaintea voastră! Noi vorbim înaintea lui Dumnezeu în Cristos; şi toate aceste lucruri le spunem, prea iubiţilor, pentru zidirea voastră. Fiindcă mă tem ca nu cumva, la venirea mea, să vă găsesc aşa cum n-aş vrea să vă găsesc şi eu însumi să fiu găsit de voi aşa cum n-aţi vrea. Mă tem să nu găsesc gîlceavă, pizma, mînii, dezbinări, vorbiri de rău, bîrfeli, îngîmfări, tulburări. Mă tem ca la venirea mea la voi, să nu mă smerească din nou Dumnezeul meu cu privire la voi, şi să trebuiască să plîng pe mulţi din cei ce au păcătuit mai înainte, şi nu s-au pocăit de necurăţiei, curvia şi spurcăciunile, pe care le-au făcut\" (2 Cor. 12:19-21).\r\n\r\nUnii dintre Corinteni începuseră să critice unele aşa zise „neajunsuri şi slăbiciuni\" ale apostolului. „Persoana\" lui Pavel era atacată în mod special. Iată cîteva din lucrurile de care îl învinuiau adversarii lui din Corint:\r\n\r\na. Laşitate şi fugă de suferinţe în faţa prigoanei. (Pavel răspunde în 11:23-33)\r\nb. Şovăială şi nestatornicie în felul în care-şi schimba mereu planurile. (Pavel răspunde în 1:15-23)\r\nc. Lipsa de scrisori de „acreditare\" din partea Bisericii din Ierusalim. (Pavel răspunde în 3-1-5-12:11-13)\r\nd. Simplitatea şi lipsa de elocvenţă a felului lui de vorbire. (Pavel răspunde în l: 12; 10:10 şi 11:6)\r\ne. Aparenta lui lipsă de personalitate. (Pavel răspunde în 10:10-13)\r\nf. Plăcerea lui de a se muta din loc în loc. (Pavel răspunde în 10:13-18)\r\ng. Poziţia lui dubioasă în ceea ce priveşte Legea lui Moise. (Pavel răspunde în cap.3 şi 4)\r\nh. Rîvna lui suspectă în adunarea de ajutoare materiale. (Pavel răspunde în cap.8; 7:2; 12:14-16)\r\ni. Zvonul despre anumite experienţe „în spirit\" care-i făceau discutabilă sănătatea mintală. (Pavel răspunde în 5:13; 12:1-10)\r\n\r\nOricît de neplăcută i-ar fi fost lui Pavel justificarea în faţa oamenilor, ea trebuia făcută deoarece la mijloc era nu numai persoana sau reputaţia lui, ci lucrarea lui în biserica din Corint şi credibilitatea lui în celelalte biserici. Eficienţa întregii lui lucrări depindea şi de felul în care îl înţelegeau oamenii.\r\n\r\nIată motivele pentru care s-a născut această a doua epistolă către cei din Corint.\r\n\r\nCuvinte cheie şi teme caracteristice: Această epistolă este o pledoarie de apărare a lui Pavel. Din ea aflăm despre necazurile apostolului şi despre felul în care-l tratau unii credincioşi din Biserici. Epistola trece însă dincolo de lămurirea poziţiei lui Pavel între credincioşii Corintului. În textul ei găsim unele dintre cele mai frumoase pasaje de teologie creştină. Iată numai cîteva dintre ele: \"făptura cea nouă\" (2 Cor. 5:17-19), „esenţa predicării creştine\" (2 Cor. 4:5, 6; 5:17-19), „vasul de lut\" (2 Cor. 4:7-18), descoperirile lui Pavel (2 Cor. 12:1-6) şi „ţepuşul\" trimis pentru smerire (2 Cor. 12:7-10).\r\n\r\nSCHIŢA CĂRŢII\r\nIntroducere (1:1, 2)\r\n\r\n1. Pavel dă socoteală de activitatea sa.\r\n(Pavel- lucrătorul)\r\n\r\n(A) în ce priveşte motivaţia sa (cap. 1 -2)\r\nPlanurile lui Pavel (1:12-2:4)\r\nIertarea celui pedepsit (2:5-13)\r\nSlujba apostolului (2:14-17)\r\n\r\n(B) în ce priveşte caracterul mesajului său (cap.3-5)\r\nUn mesaj al vieţii noi (3:1-6)\r\nUn mesaj al Noului Legămînt (3:7-18)\r\n\r\n2. Pavel îşi îndeamnă urmaşii\r\n(Pavel - părintele spiritual)\r\n\r\n(A) în ce priveşte lucrurile spirituale (cap.6-7)\r\nÎmpotriva „înjugărilor nepotrivite\" (6:14-18)\r\n\r\n(B) în ce priveşte lucrurile materiale (cap.8-9)\r\nDărnicia creştină (8:1-15)\r\n\r\n3. Pavel răspunde criticilor (Cap. 10-13)\r\n(Pavel - Apostolul)\r\n\r\n(A) criticii şi acuzaţiile lor Autoritatea apostolului (10:1-19)\r\nPurtarea lui Pavel (11:1-15)\r\nSuferinţele lui Pavel (11:16-33)\r\nDescoperirile lui Pavel (12:1-6)\r\nŢepuşul lui Pavel (12:7-10)\r\n\r\n(B) Apostolul şi dovezile apostoliei lui\r\nSemnele apostoliei (12:11-21)\r\nÎndemn la pocăinţă (13:1-10)\r\n\r\nConcluzie (13:11-14)" },
    //GALATENI
    { "GALATENI", "Explicatia Cartii: Galateni\r\n\r\n„O galateni nechibzuiţi! Cine va fermecat pe voi? Copilaşii mei, pentru care iarăşi simt durerile naşterii, pînă ce va lua Cristos chip în voi!\" - Galateni 3:1; 4:19\r\n\r\nTitulul: În originalul grec, cartea poartă numele: „Pros Galatas\" - „Către Galateni\".\r\n\r\nAutorul şi data: Prima vizită a lui Pavel în ţinuturile Galatiei a avut loc într-un timp de grea suferinţă fizica pentru apostol. Iată ce găsim scris:\r\n\r\n„Ştiţi că, în neputinţa trupului v-am propovăduit Evanghelia pentru întîia oară. Şi n-aţi arătat nici dispreţ, nici dezgust pentru ceea ce era o ispită pentru voi în trupul meu; dimpotrivă, m-aţi primit ca pe un înger al lui Dumnezeu, ca pe însuşi Cristos Isus... vă mărturisesc că, dacă ar fi fost cu putinţă, v-aţi fi scos pînă şi ochii şi mi i-aţi fi dat\" (Gal. 4:13-15).\r\n\r\nSe prea poate ca boala lui Pavel să fi fost o suferinţă cronică a ochilor care l-a făcut neplăcut la vedere. În ciuda acestui fapt, galatenii l-au îndrăgit şi probabil că apostolul şi-a petrecut convalescenţa de cîteva luni în mijlocul lor, predicîndu-le mîntuirea prin credinţa în jertfa lui Cristos.\r\n\r\nA doua vizită a lui Pavel n-a fost nici pe departe aşa de plăcută pentru apostol. Credinţa celor din Galatia se alterase şi îndreptările cerute de Pavel nu au fost primite de loc cu entuziasm:\r\n\r\n„Mă mir că treceţi aşa de repede de la Cel ce v-a chemat prin harul lui Cristos, la o altă Evanghelie...” Voi alergaţi bine, cine v-a tăiat calea ca să n-ascultaţi de adevăr? \"„M-am făcut eu oare vrăjmaşul vostru, pentru că v-am spus adevărul?\" (Gal. 1:6; 5:7; 4:16).\r\n\r\nPe fondul acestor stări, după plecarea apostolului dintre ei, apare această scrisoare (anul 56 d.Cr.). Ea este mai mult o lucrare polemică decît o scrisoare obişnuită. Este pus în discuţie felul de creştinism pe care l-au dezvoltat galatenii după ce Pavel a plecat din mijlocul lor. Un foarte potrivit început la studiul epistolei către Galateni este textul din l Cor. 3:10-15:\r\n\r\n„După harul lui Dumnezeu, care mi-a fost dat, eu, ca un meşter zidar înţelept, am pus temelia, şi un altul clădeşte deasupra. Dar fiecare să ia bine seama cum clădeşte deasupra... lucrarea fiecăruia va fi dată pe faţă...”\r\n\r\nContextul scrierii: Cînd Pavel a scris această scrisoare a fost plin de o sfîntă indignare. Inima lui de lucrător al Evangheliei era sfîşiată de durere. Ca şi în cazul Corintenilor, Pavel tremura pentru credincioşia celor convertiţi faţă de Mîntuitorul lor:\r\n\r\n„Căci sînt gelos de voi cu o gelozie după voia lui Dumnezeu, pentru că v-am logodit cu un bărbat, ca să vă înfăţişez înaintea lui Cristos ca pe o fecioară curată\". (2 Cor. 11:2)\r\n\r\nGalatia, ca şi Corintul fuseseră vizitate de „tulburători\" veniţi din Iudeea. Aceşti iudaizatori răspîndiseră pretutindeni otrava lor, răstălmăcind Evanghelia şi transformînd-o în ceva ce Domnul Isus nu a intenţionat niciodată să spună sau să facă. Atacul lor se îndrepta în două direcţii: împotriva lui Pavel însuşi şi împotriva mesajului propovăduit de el.\r\n\r\nVorbele lor sunau cam aşa: „Cine este la urma urmei acest Pavel? El nu a fost cu cei doisprăzece. S-a făcut „apostol\" el însuşi. Nu-i de mirare că mesajul lui este ciuntit, lăsînd afară părţi esenţiale ale Evangheliei. Haideţi să vă spunem noi cum stau lucrurile...”\r\n\r\nCare era mesajul iudaizatorilor? La prima vedere, ei păreau că adaugă numai cîte puţin la mesajul mîntuirii. „Credeţi în Cristos\", spuneau ei, că doar se considerau creştini, „dar să faceţi bine să vă şi tăiaţi împrejur\". Argumentul lor era că şi Pavel, la început a recomandat şi practicat acest ritual pentru cei nou convertiţi. Într-adevăr, în Fapte ni se aminteşte că după ce l-a luat cu sine pe Timotei, „l-a tăiat împrejur, din pricina iudeilor, care erau în acele locuri; căci toţi ştiau că tatăl lui era grec\" (Fapte 16:1-3).\r\n\r\nAcum, tăierea împrejur nu ar fi fost un lucru aşa de mare ca să se facă o rupere în Biserică din cauza ei, dar Pavel a văzut puţin mai departe. Dacă cei din Galatia acceptau să fie tăiaţi împrejur, aceasta nu va fi decît primul pas pe un drum fără întoarcere înspre revenirea la ţinerea întregii Legi (Gal. 5:3). Aceasta însemna „pierderea libertăţii\" (Gal.4:9), „robie spirituală\" (Gal. 5:1), părăsirea Evangheliei.\r\n\r\nEvanghelia este vestea bună despre mîntuirea dată în dar prin harul lui Dumnezeu. Dacă adaugi ceva harului, mîntuirea nu mai este gratuită, ci se capătă prin fapte.\r\n\r\nNu-i de mirare că Pavel era furios. Iată cîteva exclamaţii categorice pe care nu ne-am fi aşteptat să le găsim în gura apostolului:\r\n\r\n„Mă tem să nu mă fi ostenit degeaba pentru voi\" (Gal. 4:11).\r\n\r\n„Iată, eu, Pavel, vă spun că, dacă vă veţi tăia împrejur, Cristos nu vă va folosi la nimic\" (Gal. 5:2).\r\n\r\n„Şi schilodească-se odată cei ce vă tulbură!\" (Gal. 5:12).\r\n\r\nPentru cititorul modern s-ar putea să pară că Pavel a exagerat în reacţia lui contra învăţătorilor veniţi din Iudeea. La urma urmei, credeau şi ei în Domnul Isus şi erau fraţi cu toţi credincioşii! Dilema era însă cu mult mai adîncă. La ceasul disputei, problema era dacă noua religie a creştinilor are caracter universal sau dacă mîntuirea adusă de Cristos este numai pentru cei din neamul iudeilor. Ca să fii mîntuit, era suficient să crezi în Jertfa de la Golgota, sau trebuia mai întîi să accepţi să devii iudeu? Trebuia să accepţi şi să-ţi însuşeşti toate obiceiurile evreieşti? Trebuia să te îmbraci ca un evreu şi să împlineşti ritualurile religiei evreilor (cum predicau „iudaizatorii\" veniţi pe urmele lui Pavel)?\r\n\r\nDacă aceşti „iudaizatori\" ar fi fost lăsaţi să-şi facă jocul, probabil că noua învăţătură adusă de Domnul Isus, „Calea cea nouă\" cum o numeşte Pavel, ar fi murit de la sine în perimetrul primului secol. Dar n-a fost aşa. Istoria ni-l arată pe Pavel triumfînd. Biserica a înaintat cucerind întregul imperiu roman de atunci. Evanghelia nu a fost legată de Templu, de sacrificii sau de Legea lui Moise, lucruri despre care neamurile păgîne nu ştiau nimic şi nici nu vroiau să ştie. Prin străduinţele lui Pavel, ale lui Ştefan şi ale altor lucrători ca ei, creştinismul a ieşit din găoacea evreiască, patronată de Biserica din Ierusalim, devenind o religie transculturală cu caracter universal. Acesta i-a fost destinul trasat de însuşi întemeietorul ei:\r\n\r\n„Duceţi-vă în toată lumea, şi propovăduiţi Evanghelia la orice făptură. Cine va crede şi se va boteza, va fi mîntuit\" (Marcu 16:15).\r\n\r\nConţinutul cărţii: Scrisoarea se desfăşoară în trei mişcări distincte, fiecare acoperind cîte două capitole. Primele două capitole sînt „narative\" şi se ocupă de Pavel însuşi, autoritatea lui apostolică şi natura dumnezeiască a Evangheliei propovăduite de el. Următoarele două capitole sînt o „dispută\" privitoare la natura şi mesajul Evangheliei creştine, iar ultimele două capitole sînt „îndemnuri\" adresate direct galatenilor. Cu alte cuvinte, primele două capitole sînt „personale\", următoarele două sînt „doctrinale\", iar cele două de la urmă sînt „practice\".\r\n\r\nCuvinte cheia şi teme caracteristice:\r\n\r\n1. Poziţia lui Pavel între apostoli\r\n\r\nIudaizatorii veniţi în Galatia contestau cu tărie autoritatea lui Pavel. Atacul era îndreptat asupra calităţii lui de apostol şi asupra calităţii Evangheliei vestite de el.\r\n\r\nÎntradevăr, Pavel a fost o figură controversată în mişcarea creştină de la începutul primului secol. Faptul că el nu a fost de la început cu Domnul Isus şi că nu a fost martor al învierii Domnului îl descalificau în ochii multora ca „apostol autentic\". El nu-i însoţise pe apostoli „în toată vremea în care a trăit Domnul Isus... cu începere de la botezul lui Ioan pînă în ziuă cînd s-a înălţat la cer\". În Faptele Apostolilor l:16-26 ni se spune că Matia a fost ales „apostol\" în locul vînzătorului Iuda, exact pentru calităţile care-i lipseau lui Pavel şi astfel cercul de 12 „martori ai învierii\" (Fapte 1:22) fusese reîntregit.\r\n\r\nPe cine reprezenta atunci acest Pavel?\r\n\r\nCapitolele 1 şi 2 sînt apărarea lui Pavel împotriva acuzaţiilor aduse. În acest text găsim temelia apostolici lui Pavel şi specificul misiunii încredinţate lui de Domnul.\r\n\r\nÎn capitolul 1, Pavel îşi numeşte mesajul său „Evanghelia propovăduită de mine\"(Gal. 1:11) recunoscînd că ea se deosebeşte în unele aspecte de „Evanghelia propovăduită la Ierusalim\".\r\n\r\nÎntr-adevăr, în dezvoltarea Bisericii, Dumnezeu a hotărît ca Pavel să fie acela care să depăşească graniţele Iudaismului şi să ducă vestea mîntuirii înspre marginile pămîntului. Ucenicii Domnului s-au concentrat la început mai ales asupra „oilor pierdute ale casei lui Israel\" şi numai împotriva voinţei şi încredinţării lor au acceptat uneori să meargă la neamuri (vezi vizita lui Petru în cetatea Samariei şi vizita lui Petru în casa sutaşului Corneliu - Fapte 8 şi 10-11). A trebuit ca Dumnezeu să-l aleagă pe Pavel şi să-l trimeată ca „apostol al neamurilor\". Recrutat pe drumul Damascului şi învăţat direct de Cristosul cel înviat, probabil în pustiul Arabiei, acest Pavel a stîrnit la început tulburare oriunde şi-a propovăduit mesajul numit atît de semnificativ „Evanghelia mea\" (Rom. 2:16). Tulburarea a fost suficient de mare pentru a provoca adunarea unui Consiliu al Bisericii la Ierusalim (Fapte 15). Spre surprinderea noastră aflăm din textul care consemnează lucrările adunării din Ierusalim că în sînul Bisericii erau „unii din partida fariseilor, care crezuseră... şi care ziceau că neamurile trebuie să fie tăiate împrejur, şi să li se ceară să păzească Legea lui Moise\" (Fapte 15:5) Nu-i de mirare că înfruntarea dintre ei şi Pavel şi Barnaba a dat naştere la „multă vorbă\" (Fapte 15:7). A trebuit ca Petru şi Iacov, „care sînt priviţi ca stîlpi\" (Galat. 1:9) să ia cuvîntul şi să lămurească lucrurile. Ce au spus ei?\r\n\r\nÎn primul rînd, Petru şi-a amintit cu acest prilej ceea ce ar fi trebuit să nu uite şi anume că de fapt chemarea pe care i-o dăduse Dumnezeu fusese să facă tocmai slujba pentru care era acuzat acum Pavel:\r\n\r\n„Fraţilor, ştiţi că Dumnezeu, de o bună bucată de vreme, a făcut o alegere între voi ca, prin gura mea, neamurile să audă cuvîntul Evangheliei, şi să creadă\" (Fapte 15:7).\r\n\r\nApoi Petru le aduce aminte că în cazul lui Corneliu:\r\n\r\n„Dumnezeu, care cunoaşte inimile, a mărturisit pentru ei, şi le-a dat şi lor Duhul Sfint ca şi nouă. N-a făcut nici o deosebire între noi şi ei, întrucît le-a curăţit inimile prin credinţă\" (Fapte 15:8-9).\r\n\r\nConcluzia lui Petru şi îndemnul lui au fost:\r\n\r\n„Acum dar, de ce ispitiţi pe Dumnezeu, şi puneţi pe grumazul ucenicilor un jug, pe care nici părinţii nostrii, nici noi nu l-am putut purta? Ci credem că noi, ca şi ei, sîntem mîntuiţi prin harul Domnului Isus\" (Fapte 15:10-11).\r\n\r\nNu-i de mirare că după astfel de vorbe rostite de Petru „toată adunarea a tăcut\" şi că au ascultat cu mai multă atenţie rapoartele misionare aduse de Pavel (Fapte 15:12).\r\n\r\nUltimul care a vorbit a fost Iacov. Cuvîntarea lui este deosebit de importantă. Cu maturitatea care-l caracteriza, Iacov explică din profeţii cum Dumnezeu a hotărît o vreme în care neamurile să fie în centrul atenţiei divine în detrimentul Israelului, întărind cuvintele lui Petru, Iacov spune:\r\n\r\n„Simon a spus cum mai întîi Dumnezeu Şi-a aruncat privirile peste neamuri, ca să aleagă din mijlocul lor un popor, care să-i poarte Numele. Şi cu faptul acesta se potrivesc cuvintele proorocilor, după cum este scris: „După aceea, (după ce anume?- n.n.) Mă voi întoarce, (nu te poţi întoarce decît dacă te-ai depărtat! - n.n.), şi voi ridica din nou cortul lui David din prăbuşirea lui, îi voi zidi dărîmăturile, şi-l voi înălţa din nou; pentru ca rămăşiţa deoameni să caute pe Domnul, ca şi toate neamurile peste care este chemat numele Meu, zice Domnul, care face aceste lucruri, şi căruia îi sînt cunoscute din veşnicie\" (Fapte 15:14-18).\r\n\r\nSînt convins că foarte puţini dintre evrei se gîndiseră cum trebuie la profeţia aceasta, după cum tot foarte puţini o bagă astăzi în seamă în sînul Bisericii. Importanţa ei este dublă: ea i-a anunţat pe evrei că Dumnezeu se va întoarce pentru o vreme înspre neamuri şi anunţă neamurile că în final Dumnezeu va ridica din prăbuşirea lui, cortul lui David!\r\n\r\nConcluzia şi îndemnul lui Iacov au fost:\r\n\r\n„De aceea, eu sînt de părere să nu se mai pună greutăţi acelora dintre neamuri care se întorc la Dumnezeu...” (Fapte 15:19).\r\n\r\nCînd Pavel îşi scrie scrisoarea către galateni, problemele nu erau încă atît de clare. Ierusalimul şi „fariseii\" din Biserică îşi trimeteau încă „misionarii\" să „convertească adunările creştine născute de Pavel\". Tulburarea şi confuzia domneau pretutindeni. Iată de ce apostolul simte nevoia să se apere şi să apere adevărul mesajului Evangheliei sale:\r\n\r\n„Fraţilor, vă mărturisesc că Evanghelia propovăduită de mine nu este de obîrşie omenească; pentru că n-am primit-o, nici n-am învăţat-o de la vreun om, ci prin descoperirea lui Isus Cristos...”(Galateni 1:11-23)\r\n\r\nPavel mărturiseşte că, deşi independent de apostolii de la Ierusalim, el n-a lucrat fără cunoştinţa şi încuviinţarea lor (Galateni 2:1-9). Această recunoaştere a Evangheliei lui Pavel fusese pusă la îndoială de atitudinea lui Petru în Antiohia:\r\n\r\n„Dar cînd a venit Petru în Antiohia, i-am stătut împotrivă în faţă, căci era de osîndit\" (Galateni 3:1). Deosebirea pe care a făcut-o Petru între creştinii dintre evrei şi creştinii dintre neamuri trebuia osîndită pe faţă şi Pavel nu a ezitat să o facă. Era în joc mîntuirea mulţimilor care primiseră credinţa:\r\n\r\n„Nu vreau să fac zadarnic harul lui Dumnezeu; căci dacă neprihănirea se capătă prin Lege, degeaba a murit Cristos\" (Galateni 2:21).\r\n\r\nPentru iudaizatori, ca să fii un bun creştin trebuia mai întîi să crezi în Cristos şi apoi să ajungi să împlineşti toată Legea. Cu alte cuvinte, pentru a deveni pe deplin creştin, trebuia să devii de asemenea „iudeu\"; să accepţi tăierea împrejur, Sabatul şi toate celelalte ceremonii iudaice. Pentru Pavel, aşa ceva era exact contrariul creştinismului. Mîntuirea s-ar fi putut obţine atunci, nu prin har, ci prin ceea ce putea face un om, printr-un semn în carne şi prin abilitatea cuiva de a ţine Legea. Pavel ştia însă că mîntuirea se primea prin har şi numai prin credinţă. Toate strădaniile omului nu puteau duce nicăieri. Ceea ce era necesar era nu împlinirea Legii, ci declararea unui faliment total, abandonarea de sine la picioarele crucii lui Cristos şi aruncarea în braţele iubitoare ale Mîntuitorului. iudeul era înclinat să spună: „Doamne uite lucrările pe care le-am făcut. Iată semnul tăierii mele împrejur. Dă-mi acum mîntuirea pe care mi-am cîştigat-o.\" Pavel privea aşa ceva drept o blasfemie la adresa sacrificiului ispăşitor al lui Cristos.\r\n\r\n„Dar, spuneau iudeii, cel mai mare lucru din viaţa noastră ca popor al lui Dumnezeu este Legea dată nouă prin Moise. Fără ea n-am fi ştiut ce înseamnă să trăieşti după placul lui Dumnezeu. Cum să renunţăm acum la ea? Cum să renunţăm la trecutul nostru de popor al Domnului?\"\r\n\r\n„Foarte bine, răspundea Pavel în capitolul 3:1-29, să vedem atunci cine este strămoşul nostru: Moise sau Avraam? Avraam cu siguranţă. Şi cum a căpătat Avraam trecere înaintea lui Dumnezeu, prin faptele Legii sau prin credinţă?\r\n\r\n„Avraam a crezut pe Dumnezeu, şi credinţa aceasta i-a fost socotită neprihănire\" (Gal. 3:6)\r\n\r\n„Scriptura, de asemenea, fiindcă prevedea că Dumnezeu va socoti neprihănite pe neamuri, prin credinţă, a vestit mai dinainte lui Avraam această veste bună: „Toate neamurile vor fi binecuvîntate în tine.\" Şi că nimeni nu este socotit neprihănit prin Lege, este învederat, căci „cel neprihănit va trăi prin credinţă”.\r\n\r\nÎnsă Legea nu se întemeiază pe credinţă; ci ea zice: „Cine va face aceste lucruri, va trăi prin ele\". Cristos ne-a răscumpărat din blestemul Legii... pentru ca binecuvîntarea vestită lui Avraam să vină peste neamuri, în Cristos Isus\" (Gal. 3:8-14)\r\n\r\nÎntr-adevăr, „sămînţa\" lui Avraam este acest Isus Cristos în care sînt binecuvîntaţi toţi credincioşii (Gal. 3:16).\r\n\r\nPentru Pavel, Legea a funcţionat ca un „pedagog spre Cristos\" (Gal. 3:24). Pedagogul era pe atunci sluga ce ducea elevul la şcoala maestrului său. Rolul pedagogului înceta odată cu apariţia profesorului:\r\n\r\n„După ce a venit credinţa, nu mai sîntem sub îndrumătorul acesta... Nu mai este nici iudeu, nici grec, nu mai este nici rob, nici slobod... fiindcă toţi sînteţi una în Cristos Isus. Şi dacă sînteţi ai lui Cristos, sînteţi „sămînţa\" lui Avraam, moştenitori prin făgăduinţă\" (Gal. 3:25-29).\r\n\r\nSe poate spune că Iudaismul a fost „leagănul creştinismului\" şi noi am putea adăuga „şi era gata, gata să-i fie şi mormîntul!\"\r\n\r\nA trebuit ca Dumnezeu să-l ridice pe Pavel, acest Moise al Bisericii, pentru ca prin el să fim eliberaţi din robia „învăţăturilor începătoare\" ale Legii. Cu entuziasmul cu care Moise a ridicat înaintea poporului tablele Legii, Pavel ridică steagul Crucii lui Cristos. Moise venise să ne dea „mărturia\" şi ne-a făcut „robi ai păcatului care clocoteşte în noi\". Pavel ne prezintă un Cristos al Crucii care a venit să ne facă „cu adevărat slobozi\" (Ioan 8:36). El are „toată puterea în cer şi pe pămînt\" şi poate face acest lucru!\r\n\r\n2. Roada Duhului Sfînt\r\n\r\nPavel scrie unei colectivităţi de oameni care se ocupau cu agricultura. Din această pricină el îşi alege termeni corespunzători, dovedind o mare flexibilitate în exprimare şi o deosebită pricepere în adaptarea mesajului la puterile şi vocabularul ascultătorilor. Vorbind despre rezultatele produse de lucrarea Crucii în viaţa celor credincioşi, el le numeşte „ROADA\" şi le prezintă în contrast cu „faptele firii pămînteşti\" (Gal. 5:16-26). Cele 17 fapte ale firii sînt contrastate cu cele 9 rodiri ale Duhului. Să nu credeţi cumva că această exprimare simplă este şi simplistă! Roada Duhului, aşa cum o prezintă aici Pavel, este aşezată în trei grupe distincte de cîte trei şi cuprinde într-o minunată aşezare transformarea totală pe care o face Duhul lui Dumnezeu atunci cînd pătrunde în viaţa cuiva.\r\n\r\nRodire către Dumnezeu: „Dragostea, bucuria, pacea,\r\nRodire către alţii: Îndelunga răbdare, bunătatea, facerea de bine,\r\nRodire faţă de noi înşine: credincioşia, blîndeţea, înfrinarea poftelor.”\r\n\r\nConcluzia acestei prezentări este cuprinsă în Gal. 6:7-9. Acest pasaj ne arată că nu-l putem „duce\" pe Dumnezeu. Fiecare îşi va primi răsplata după alegerea pe care a făcut-o:\r\n\r\n„Cine seamănă în firea pămîntească, va secera din firea pâmîntească putrezirea; dar cine seamănă în Duhul, va secera din Duhul viaţa veşnică\" (Gal. 6:8).\r\n\r\nSecerişul nu se va face după cît de mult am ştiut, ci după cît de mult am semănat!\r\n\r\n3. „Semnele Domnului Isus Cristos\"\r\n\r\nAcestea sînt semne pe care le purta Pavel pe trupul lui şi pentru care îl batjocoreau unii (Gal. 6:17). Cuvîntul grec folosit aici este„stigmata\"şi se poate traduce prin: sigiliu, semnul de proprietate aşezat uneori pe spatele, pe faţa sau pe braţul unui sclav şi pe pielea unor animale.\r\n\r\nCare să fi fost „stigmata\" lui Pavel?\r\n\r\nErau semnele bătăilor şi loviturilor primite pentru mărturia lui creştină!\r\n\r\n„...arătăm că sîntem vrednici slujitori”...„în lovituri fără număr... De cinci ori am căpătat de la iudei patruzeci de lovituri fără una; de trei ori am fost împroşcat cu pietre; de trei ori s-a sfărîmat corabia cu mine...\" (2 Cor. 6:4; 11:23-25)\r\n\r\nMîinile bătătorite ale unui lucrător îi arată osteneală, cicatricile lui Pavel, pentru care unii îl puneau în rînd cu „tulburătorii\" şi cu făcătorii de rele, dovedeau credincioşia lui în slujirea creştină. Faţa arsă de soare a căpitanului de marină, rănile unui soldat şi ridurile de pe fruntea unei mame nu sînt semne de dispreţuit. „Semnele Domnului Isus\" purtate de Pavel pe trupul lui nu erau temei de batjocură! „Iudaizatorii\" aveau scrisori de acreditare de la Ierusalim, Pavel purta pe trupul său semnele unei acreditări mult mai înalte. Ce suferinţe înduraseră „iudaizatorii\" pentru Evanghelie? Niciuna. Ei au ştiut să-şi păzească pielea. Pavel în schimb devenise „dizgraţios\" la privit pentru că se făcuse una cu „Omul durerii\"!\r\n\r\nCUPRINSUL CĂRŢII\r\n\"Eliberare prin Evanghelie\"\r\n\r\nCuvînt de salut (1:1-5)\r\n\r\n1. Autenticitatea Evangheliei lui Pavel (1 şi 2)\r\nVeritabilă în ce priveşte originea ei (cap. 1)\r\nVeritabilă în ce priveşte natura ei (cap. 2)\r\n\r\n2. Superioritatea Evangheliei creştine (3 şi 4)\r\nÎn noile relaţii pe care le produce (cap.3)\r\nÎn privilegiile pe care le aduce (cap. 4)\r\n\r\n3. Adevărata slobozenie prin Evanghelie (5 şi 6)\r\nSlujba iubirii pune capăt robiei Legii (5:1-15)\r\nDuhul pune capăt robiei în firea pămăntească (5:16-6:10)\r\nCuvînt de încheiere (6:11-18)" },
    //EFESENI
    { "EFESENI", "Explicatia Cartii: Efeseni\r\n\r\nAşa cum s-a putut observa deja, Romani şi Galateni sînt „visteriile\" în care se găsesc comorile de adevăr despre mîntuirea personală prin credinţa în Crucea Domnului Isus Cristos. Ele au adus lumii mesajul mîntuirii prin credinţă:\r\n\r\n„Dar acum s-a arătat o altă neprihănire, pe care o dă Dumnezeu, fără Lege... şi anume, neprihănirea dată de Dumnezeu, care vine prin credinţa în Cristos, pentru toţi şi peste toţi cei ce cred în El. Nu este nici o deosebire\" (Rom. 3:21-22).\r\n\r\nUrmătoarele trei epistole de care ne vom ocupa în continuare, Efeseni, Filipeni şi Coloseni, ne pun în faţa unui tablou care depăşeşte cadrul mîntuirii personale, ajungînd să ne descopere, dincolo de poziţia noastră în Cristos şi a lui Cristos în noi, poziţia pe care o ocupă Domnul Isus Cristos în economia planului pe care-l are Dumnezeu cu lumea.\r\n\r\nIdeea centrală a lor este prezentarea lui Cristos drept „Cap\" al întregii creştinătăţi care-i formează „trupul\".\r\n\r\nTitlul: Numirea „Pro Efesious\" - „Către Efeseni\" nu există în toate copiile găsite.\r\n\r\nAutorul: Efeseni, Coloseni, Filimon şi Filipeni sînt supranumite şi „epistolele captivităţii\", deoarece au fost scrise de Pavel în timpul cînd se afla în închisoare.\r\n\r\nData: În timpul primei detenţii în închisoarea Romei (60-62 d.Cr.)\r\n\r\nContextul scrierii: Pavel se afla din nou închis. De data aceasta nu la Cezareea, ci la Roma. Apostolul se numeşte pe sine: „întemniţatul lui Isus Cristos pentru voi\" (Efeseni 3:1). El îşi sfătuieşte cititorii cu autoritatea unuia care este „întemniţat pentru Domnul\" (Efes. 4:1), devenind pentru Evanghelie „un sol în lanţuri\" (Efes. 6:20). Epafra vine să-l viziteze şi-i aduce veşti nu prea bune despre Biserica din Colose. Pavel tocmai vroia să-l trimită la Colose pe Onisim, sclavul fugit de la stăpîn şi acum „încreştinat\" de apostol. Profitînd de această „potrivire\", Pavel se aşterne iarăşi la masa de scris şi le adresează Colosenilor o scrisoare în care tratează „rătăcirea\" care apăruse în această Biserică. În timpul scrierii, apostolul îşi dă seama însă că subiectul este mult prea important pentru a nu fi împărtăşit şi celorlalte Biserici. De aici se naşte hotărîrea lui de a dezvolta tema abordata în scrisoarea către Colose, amplificînd-o într-un fel de „scrisoare circulară\" către toate bisericile. Iniţial, această scrisoare le-a fost adresată celor din Laodicea. Intenţia lui Pavel a fost ca între biserici să se facă apoi un schimb de scrisori, astfel ca toţi să le citească pe toate:\r\n\r\n„După ce va fi citită această epistolă la voi, faceţi aşa ca să fie citită şi în Biserica Laodicienilor; şi voi, la rîndul vostru, să citiţi epistola care vă va veni din Laodicea\" (Coloseni 4:16).\r\n\r\nManuscrisele cele mai timpurii nu păstrează precizarea: „către Efeseni\", iar noi sîntem îndreptăţiţi să credem că această epistolă este de fapt o copie a acelei scrisori adresate celor din Laodicea. Epistola ar putea purta de fapt orice nume ca destinaţie, căci ea nu se adresează unei situaţii locale, ci prezintă învăţături pe care trebuie să le cunoască Bisericile de pretutindeni şi din totdeauna.\r\n\r\nFilimon, Coloseni şi Efeseni pleacă deci împreună din mîna lui Pavel într-un timp în care acesta se găsea în închisoare. Scrisorile sînt purtate probabil de Tihic despre care găsim scris şi în Efeseni şi-n Coloseni (Col. 4:7-9; Efeseni 6:21).\r\n\r\nConţinutul cărţii: Efeseni este cea mai profundă scriere a lui Pavel, „regina tuturor epistolelor\" lui. Efeseni şi Coloseni seamănă foarte mult în limbaj una cu alta. Nu mai puţin de 75 de versete din cele 155 de versete ale scrisorii către cei din Efes se găsesc aproximativ identic în scrisoarea adresată Colosenilor. Şi acest fapt întăreşte convingerea noastră că, plecînd de la o problemă locală, apostolul a dezvoltat-o apoi într-o lucrare adresată tuturor Bisericilor.\r\n\r\nTema epistolei către Efeseni este strîns legată de tema epistolei scrisă Colosenilor. Gîndul central din scrisoarea către cei din Colose este „atotsuficienţa lui Isus Cristos\".\r\n\r\nÎn Isus Cristos locuiesc „toate comorile înţelepciunii şi ale ştiinţei\". (Col. 2:3) „Căci Dumnezeu a vrut ca toată plinătateasă locuiască în El\". (Col. 1:19) „Căci în El locuieşte trupeşte toată plinătatea Dumnezeirii\". (Col. 2:9) El este singurul prin „care avem răscumpărarea, prin sîngele Lui, iertarea păcatelor\". (Col. 1:14) Toată epistola adresată Colosenilor este clădită pe ideea că Cristos este „atotsuficient\" pentru toate nevoile noastre.\r\n\r\nEpistola scrisă Efesenilor este o dezvoltarea a acestei idei. Întreaga epistolă poate fi rezumată în două versete plasate în primul capitol:\r\n\r\n„...căci a binevoit să ne descopere taina voiei Sale, după planul pe care-l alcătuise în Sine însuşi, ca să-l aducă la îndeplinire la plinirea vremurilor, spre a-Şi uni iarăşi într-unul, în Cristos, toate lucrurile; cele din ceruri şi cele de pe pămînt\" (Efes. 1: 9-10).\r\n\r\nMesajul central al epistolei este: refacerea acelui „Unu\" cosmic dinaintea căderii; redobîndirea unităţii prin Cristos; unirea cerului cu pămîntul prin poziţia şi lucrarea lui Isus Cristos.\r\n\r\nCristos este în acelaşi timp şi centrul în care trebuie să se întîlnească toate lucrurile şi substanţa care le leagă pe toate împreună. Toate sînt unite „în El\", dar şi „prin El\". Există în lume o „rupere\" care a adus durere şi suferinţă. Păcatul neascultării i-a separat pe cei vinovaţi de Dumnezeu producînd o stare de tensiune, un dezechilibru în tot ceea ce există. Scrisoarea către Efeseni ne ajută să ne dăm seamă mai întîi de existenţa acestei stări de conflict care este în natură, în om, între oameni, în timp, în eternitate, între neamuri, între fiinţele cereşti, între om şi Dumnezeu şi ne îndreaptă apoi spre convingerea că această stare se poate îndepărta numai atunci cînd toate lucrurile, toate puterile şi toţi oamenii se vor uni „în Cristos\". Sarcina supremă a Bisericii, aşa cum o vede Pavel, este să spună tuturor celor vii că unitatea dintre toţi oamenii şi dintre oameni şi lumea în care trăim, unitatea după care tînjim cu toţii în familie şi în societate nu poate fi realizată în afara lui Cristos. Mesajul Bisericii trebuie însoţit de o demonstraţie practică a realizării unei vieţi noi în perimetrul celor ce l-au primit pe Cristos. Iată de ce Stott numeşte comentariul său în epistola către Efeseni: „Societatea cea nouă a lui Dumnezeu\".\r\n\r\nChiar şi un cititor superficial observă imediat că cele 6 capitole ale scrisorii sînt împărţite în două jumătăţi distincte. „Doxologia\" de la sfîrşitul capitolului 3 este un fel de piatră de hotar între două secţiuni clar deosebite.\r\n\r\nPrima parte a epistolei, cuprinzînd capitolele de la 1 la 3, este doctrinală, iar cea de a doua jumătate, aşezată în capitolele 4, 5 şi 6 este practică.\r\n\r\nCuvinte cheie şi teme caracteristice: Prima jumătate a epistolei se ocupă cu dezvăluirea „poziţiei\" noastre în Cristos, iar a doua jumătate ne arată „practica\" în care trebuie să se evidenţieze această poziţie. În primele capitole vedem ceea ce a cîştigat Cristos pentru noi la Golgota, iar în cele de la urmă vedem ce fel de viaţă vrea El să desfăşoare în noi şi prin intermediul nostru. Fiecare din cele două secţiuni debutează cu cîte un verset care-i anunţă tema:\r\n\r\nDespre POZIŢIA noastră în Cristos: „Binecuvîntat să fie Dumnezeu, Tatăl Domnului nostru Isus Cristos, care ne-a binecuvîntat cu tot felul de binecuvîntări duhovniceşti, în locurile cereşti, în Cristos\" (Efes. l:3).\r\n\r\nDespre PRACTICA pe care trebuie s-o manifestăm: „Vă sfătuiesc dar eu, cel întemniţat pentru Domnul, să vă purtaţi într-un chip vrednic de chemarea, pe care a-ţi primit-o...” (Efes. 4:1)\r\n\r\nCaracteristic epistolei către Efeseni este şi „Taina ascunsă de veacuri în Dumnezeu\" despre care vorbeşte Pavel în cap. 3:4-10. Ea este o învăţătură despre care a vorbit Domnul Isus, dar nu a explicat-o în timpul vieţii Sale pămînteşti. Proclamarea acestei învăţături a fost misiunea specifică pentru care l-a ales Dumnezeu pe apostolul Pavel. Pusă în cuvinte foarte simple „taina\" descoperită este vestea că Cristos, în loc să preia imediat după venirea pe pămînt „domnia\" în împărăţia terestră pe care o aşteptau evreii, va accepta dimpotrivă să fie lepădat, să sufere, să moară pe o cruce şi după înviere şi înălţare, să dispară complet de pe scena lumii pămîntene. Taină fusese că, după ce se va aşeza la dreapta Tatălui, primind o autoritate şi o putere superioară tuturor creaturilor din veacul acesta sau din veacul viitor, El va strînge la Sine un număr de „aleşi\" indiferent din ce neam fac parte, pe care îi va duce într-o intimitate atît de mare cu Sine însuşi, încît ei nu vor mai putea fi numiţi cu alţi termeni decît aceia de: „trupul\" Său, „mireasa\" Sa şi „Templul\" Său. Aceste trei metafore definesc, fiecare, unitatea credincioşilor cu Cristos în viaţă („trupul\"), în dragoste („mireasa\") şi în slavă („templul ).\r\n\r\nImaginea luptătorului îmbrăcat cu „toată armătura lui Dumnezeu\" din finalul epistolei este un tablou cum nu se poate mai nimerit pentru încheierea epistolei. În faţa a tot ceea ce a făcut Dumnezeu pentru noi credincioşii, nu trebuie să ne culcăm pe perna moale a delăsării, ci trebuie să veghem şi să ne apărăm împotriva atacurilor Diavolului.\r\n\r\nApostolul Pavel ne spune tuturor: „Întăriţi-vă în Domnul şi în puterea tăriei Lui\" (Efeseni 6:10).\r\n\r\nCUPRINSUL CĂRŢII\r\nCuvînt de salut (cap.1:1, 2)\r\n\r\n1. Poziţia noastră în Cristos (cap. 1-3)\r\na. Recunoştinţă pentru binecuvîntările duhovniceşti (1:3-14)\r\nb. Rugăciune pentru pătrundere duhovnicească a lucrurilor (1:15-23)\r\nc. O nouă poziţie în Cristos (2:1-10)\r\nd. O nouă relaţie prin Cristos (2:11-22)\r\ne. Priceperea Tainei divine (3:1-12)\r\nf. Primirea Plinătăţii divine (3:13-21)\r\n\r\n2. Practica unei vieţii în Cristos (cap. 4-6)\r\nÎn viaţa Bisericii (4:1-16)\r\nÎn viaţa fiecărui credincios (4:17-32)\r\n\r\nViaţa creştină este o viaţă:\r\na. trăită în dragoste (5:1-20)\r\nb. trăită separat de imoralitate (5:3-7)\r\nc. trăită în lumină (5:8-14)\r\nd. trăită în înţelepciune (5:15-20)\r\ne. trăită în supunere reciprocă (5:21-6:9)\r\nf. trăită în biruinţă (6:10-20)\r\n\r\nCuvînt de încheiere (6:21-24)" },
    //FILIPENI
    { "FILIPENI", "Explicatia Cartii: Filipeni\r\n\r\nTitlul: În originalul grec, cartea poartă numele: „Pros Filippesious\" - „Către Filipeni\".\r\n\r\nAutorul: Pavel a sosit pentru prima dată în Filipi prin preajma anului 52 d.Cr. Începuturile Bisericii din Filipi sînt relatate de Luca în capitolul 16 din Faptele Apostolilor. El ne spune că Pavel a avut o vedenie în care un om din Macedonia i-a spus: „Treci în Macedonia şi ajută-ne!\" (Fapte 16:9). Ca răspuns la vedenie, apostolul a luat corabia din Troa pentru Neapolis („oraşul nou\") de pe continentul Europei şi s-a îndreptat repede spre cetatea Filipi. Vizita şi lucrarea lui Pavel în cetate s-au centrat în jurul a trei persoane: Lidia, vînzătoarea de purpură, roaba cu un duh de ghicire pe care a eliberat-o Pavel şi temnicerul cetăţii. Lucrarea lui Cristos în Filipi a început deci cu o femeie din Asia (Tiatira), cu o sclavă din Grecia şi cu un Roman. Acestora li s-au adăugat cei cîţiva evrei reprezentaţi de „femeile\" adunate la rugăciune pe malul apei. Împreună, aceştia au format nucleul noii Biserici (Fapte 16:40).\r\n\r\nData: Epistola către Filipeni a fost scrisă în anul 62 d.Cr., cam la treizeci de ani după înălţarea Domnului Isus la cer şi cam la zece ani după ce apostolul Pavel predicase pentru prima oară în Filipi.\r\n\r\nContextul scrierii: Creştinismul era încă tînăr, plin de putere şi de entuziasm. Biserica apăruse ca o revărsare de prospeţime într-o lume intrată în putrefacţia păcatului. Totul în jur părea că se prăbuşeşte. Religiile de altădată nu mai erau decît umbre. Filosofiile se transformaseră în păreri discutabile. Deasupra tuturor se aşezase, strivind tot, sandala imperială a Romei. Forţa subjugase încă o dată spiritul. Carnea şi pornirile ei puseseră stăpînirea pe lume. Egoismul, cruzimea, jaful şi uciderea erau răspîndite la culme pe toată faţa pămîntului.\r\n\r\n„Ce este Adevărul?\" nu mai preocupa aproape pe nimeni. „Adevărul\" zilei era dictat de cei care deţineau puterea. Religiozitatea, evlavia şi înfrînarea nu se mai întîlneau decît pe alocuri. „Omul\" era în centrul preocupărilor lumii. Cezarul fusese declarat „Zeu\" şi astfel de zei dictau destinul popoarelor. În confuzia care domnea pretutindeni, Evanghelia creştină a apărut ca un fulger în noapte, brăzdînd întunerecul care se lăsase greu. Ea se prezenta tuturor ca o veste nouă venită din ceruri şi într-adevăr de aşa ceva aveau nevoie oamenii. Învăţătura ei antagoniza pretenţiile Romei care-l ridicase pe Cezar printre zei. Evanghelia prezenta calea inversă, calea prin care Dumnezeu însuşi coborîse în mijlocul oamenilor. Biserica creştină prezenta lumii, nu o nouă învăţătură şi nici o nouă formă de organizare, ci o Persoană care a trăit o viaţă exemplară. Fusese o viaţă trăită între oameni, dar într-atît de deosebită, într-atît de unică în măreţia ei, în iubirea ei ne-pămînteană, în absoluta ei puritate şi-n completa ei dăruire de sine, încît această viaţă atrăgea cu puterea ei de fascinaţie pe toţi cei care aflau de existenţa ei. Isus-Mîntuitorul, arătase lumii pentru prima dată ce face şi cum este Dumnezeul care a creiat universul.\r\n\r\nViaţa acestui întemeietor al creştinismului nu se sfîrşise odată cu moartea. Ea continua să trăiască dincolo de înviere, manifestîndu-se cu putere în existenţa şi faptele Bisericii. Cristos trăia şi în cer, dar şi în creştinii primului secol.\r\n\r\nFilipi era un fel de Romă în miniatură. Limba oficială era latina, dar pretutindeni pe străzi se vorbea greceşte. Numele vechi al cetăţii Filipi fusese mai întîi Datos, iar apoi Krenides, care se poate traduce prin „izvoarele\" sau „puţurile\". Numele de „Filipi\" îi fusese dat de tatăl lui Alexandru cel Mare, Filip, care îşi făcuse o avere imensă cu aurul scos din minele din împrejurimi. Cetatea se afla într-o regiune extraordinar de bogată. Un sol fertil şi un subsol încărcat de metale preţioase îi aduseseră foarte curînd celebritatea. Dincolo de toate acestea, plasarea oraşului într-o depresiune montană care funcţiona ca o veritabilă „poartă\" între Europa şi Asia, îi adusese afluenţă comercială şi-l determinase pe Cezar Augustus să declare cetatea Filipi: „Colonie Romană\". Coloniile din Imperiu nu erau cuceriri noi, ca în terminologiile folosite de noi astăzi. Ele erau mai degrabă oraşe noi, formate din „Legionari\" retraşi la pensie după împlinirea datoriilor militare. Aşa se explică faptul că, în ciuda numărului de locuitori greco-macedoneni, majoritatea cetăţenilor din Filipi se numeau „romani\":\r\n\r\n„I-au dat pe mîna dregătorilor şi au zis: „Oamenii aceştia ne tulbură cetatea; sînt nişte ludei, care vestesc nişte obiceiuri pe care noi, Romanii, nu trebuie să le primim, nici să le urmăm\" (Fapte 16:20-21).\r\n\r\nÎn Filipi existau puţini iudei. Fără îndoială că pe aceştia îi adusese acolo, nu caracterul militar al oraşului, ci comerţul înfloritor care-i trecea porţile. Numărul mic explică de ce iudeii nu aveau o Sinagogă acolo, ci doar „un loc de rugăciune\", aşezat lîngă un curs de apă, unde se puteau îndeplini spălările rituale evreieşti (Fapte 16:13).\r\n\r\nAcest grup de credincioşi a devenit imediat cea mai scumpă adunare pentru inima lui Pavel. Relaţiile dintre apostol şi cei din Filipi nu au fost niciodată tulburate de neîncredere sau de păcate ascunse, aşa cum a fost cazul cu alte Biserici înfiinţate de Pavel. „Din cea dintîi zi\" a existenţei lor, cei din Filipi s-au alăturat eforturilor lui Pavel pentru înaintarea Evangheliei în lume:\r\n\r\n„Mulţumesc Dumnezeului meu pentru toată aducerea aminte pe care o păstrez pentru voi. În toate rugăciunile mele mă rog pentru voi, cu bucurie pentru partea, pe care o luaţi la Evanghelie, din cea dintîi zi pînă acum... Întrucît, atît în lanţurile mele, cît şi în apărarea şi întărirea Evangheliei, voi sînteţi toţi părtaşi aceluiaşi har\" (Filipeni 1:3-9).\r\n\r\nCredincioşii din Filipi s-au identificat cu strădaniile şi suferinţele apostolului, trimiţîndu-i suportul lor material ori de cîte ori au avut ocazia. Cel puţin de două ori i-au trimis lui Pavel ajutor în Tesalonic (Filipeni 4:16). Apoi cînd apostolul a părăsit Macedonia i-au trimis din nou (Filipeni 4:15) şi i-au trimis iarăşi cînd avea nevoie de ajutor în sudul Greciei, acolo unde apostolul s-a ferit să primească ceva de la Corintenii care, cel puţin în privinţa dărniciei, aveau o cu totul altfel de inimă (2 Cor. 11:9).\r\n\r\nDin epistola lui Pavel către Filipeni aflăm că grija lor pentru apostol a continuat să se manifeste şi după „întemniţarea\" lui în Roma (Filip. 4:10-19).\r\n\r\nConţinutul cărţii: Această scurtă epistolă nu este o lucrare de lămuriri doctrinare, ci pur şi simplu o „scrisoare\". Comentatorii o numesc: „Cea mai puţin dogmatică dintre toate epistolele lui Pavel\". Caracterul ei este practic, nu profesoral; corectiv, nu informativ; o scrisoare de dragoste, de recomandare şi de mulţumire. Desigur, tot textul ei este impregnat cu învăţătură creştină, dar aceasta apare numai ca un fundal pentru desfăşurarea ideilor şi îndemnurilor apostolului. Problemele din Filipi nu ajunseseră de fapt să fie „probleme\": doar puţină mîndrie, puţină lipsă de unitate, puţină ceartă, puţin murmur şi un început de amărăciune. Cele patru capitole ale cărţii sînt scrise de Pavel pentru a atinge cîteva scopuri:\r\n\r\na. Epistola este o scrisoare de mulţumire. Trecuseră aproximativ 10 ani de cînd se cunoştea cu cei din Filipi şi aceştia tot mai continuau să-l mai trimită suport material.\r\n\r\nb. Epistola este o informare despre situaţia lui Epafrodit, trimisul filipenilor pe lîngă Pavel. Se pare că cei din Filipi l-au trimis pe acest Epafrodit la Pavel nu numai ca să-i ducă banii, ci şi ca să rămînă cu el şi să-i slujească. Din păcate, Epafrodit s-a îmbolnăvit însă foarte serios. Cînd s-a vindecat, l-a cuprins dorul de cei de acasă şi mai era şi îngrijorat la gîndul că cei din Filipi aflaseră de suferinţa lui. Apostolul Pavel s-a hotărît să-l trimeată acum înapoi. Nevrînd însă ca unii din Filipi să-l considere pe Epafrodit ca pe unul care şi-a părăsit slujirea, Pavel îl trimite împreună cu o scrisoare de recunoştinţă în care-l laudă puţin: „Primiţi-l deci în Domnul cu toată bucuria; şi preţuiţi pe astfel de oameni. Căci pentru lucrul lui Cristos a fost el aproape de moarte, şi şi-a pus viaţa în joc, ca să împlinească ce lipsea slujbei voastre pentru mine\" (Filip. 2:29-30).\r\n\r\nc. Epistola este un îndemn la trăirea în unitate. Pavel aflase că două femei din adunarea din Filipi creiaseră o stare de tensiune între fraţi: „Deci, dacă este vreo îndemnare în Cristos... faceţi-mi bucuria deplină şi aveţi o simţire, o dragoste, un suflet şi un gînd... Îndemn pe Evodia şi îndemn pe Sintichia să fie cu un gînd în Domnul. Şi pe tine, adevărat tovarăş de jug, te rog să vii în ajutorul femeilor acestora, care au lucrat împreună cu mine pentru Evanghelie...” (Filip.4:l-3).\r\n\r\nd. Epistola este un mesaj de încurajare în faţa valului de persecuţie care se pornise peste creştini. De fapt, apostolul Pavel a părăsit el însuşi cetatea Filipi după ce a fost bătut pe nedrept cu nuiele şi pus în mod abuziv şi ilegal în închisoare. Aruncat împreună cu Sila în „temniţa din lăuntru, cu picioarele în butuci\" el a fost folosit în mod minunat de Dumnezeu pentru „Evanghelizarea\" deţinuţilor şi pentru convertirea temnicerului. Am putea spune deci că cei din Filipi au moştenit de la Pavel persecuţia:\r\n\r\n„Căci ai privire la Cristos, vouă vi s-a dat harul nu numai să credeţi în El, ci să şi pătimiţi pentru El... fără să vă lăsaţi înspăimîntaţi de potrivnici; lucrul acesta va fi pentru ei o dovadă de pierzare, şi de mîntuirea voastră, şi aceasta de la Dumnezeu\" (Filip, l:29, 28).\r\n\r\ne. Epistola este un avertisment împotriva infiltrărilor \"iudaizatorilor\": „Păziţi-vă de cîinii aceia; păziţi-vă de lucrătorii aceia răi; păziţi-vă de scrijilaţii aceia! „Căci cei tăiaţi împrejur sîntem noi, care slujim lui Dumnezeu, prin Duhul lui Dumnezeu, care ne lăudăm în Cristos Isus, şi care nu ne punem nădejdea în lucrurile pămînteşti\" (Filip. 3:2, 3).\r\n\r\nCuvinte cheie şi teme caracteristice: Există anumite idei majore care se repetă în epistolă:\r\n\r\n1. Ideea bucuriei creştine (1:4, 18, 25; 2:16, 17, 18, 28; 3:1, 3; 4:1, 4). Bucuria despre care vorbeşte Pavel în această epistolă este „bucuria-roadă a Duhului Sfînt\", care caracterizează toate momentele vieţii creştine. Ea poate fi experimentată indiferent de vitregia împrejurărilor sau de momentele de suferinţă în firea pămîntească. Pavel şi Sila s-au rugat şi au cîntat cîntări de laudă chiar dacă fuseseră bătuţi crunt cu nuielele şi le erau picioarele puse în butuci (Fapte 16:25).\r\n\r\nUn comentator din secolul trei spune că întregul mesaj al epistolei ar putea fi rezumat în aceste cuvinte: „Eu, Pavel, mă bucur! Voi vă bucuraţi?\"\r\n\r\n2. Ideea \"cîştigului în Cristos\". Lucruri contradictorii care răstoarnă toate socotelile, descoperind creştinilor orizonturi noi care trebuiesc explorate (1:21, 23; 3:7, 8, 14; 4:19).\r\n\r\nCUPRINSUL CĂRŢII\r\nIntroducere (1:1 -2)\r\n\r\nI. Scurtă dare de seamă\r\na. Rugăciune pentru filipeni, 1:3-11\r\nb. Pavel suferă pentru Evanghelie, 1:12-26\r\nc. Filipenii sufăr pentru Evanghelie, 1:27-30\r\n\r\nII. Apel la smerenie 2:1-30\r\na. Dorinţa lui Pavel, 2:1-4\r\nb. Pilda lui Cristos, 2:5-16\r\nc. Pilda lui Pavel, 2:17-18\r\nd. Pilda lui Timotei, 2:19-24\r\ne. Pilda lui Epafrodit, 2:25-30\r\n\r\nIII. Apel la cunoaştere, 3:1-21\r\na. Avertisment împotriva ereziei, 3:1-9\r\nb. Pilda alergării lui Pavel, 3:10-17\r\nc. Avertisment împotriva trăirii în pacat, 3:17-21\r\n\r\nIV. Apel la pace, 4:1-23\r\na. Pace între fraţi, 4:1-3\r\nb. Pace cu Domnul, 4:4-9\r\nc. Pace în orice situaţie, 4:10-19\r\n\r\nÎncheiere, 4:20-23" },
    //COLOSENI
    { "COLOSENI", "Explicatia Cartii: Coloseni\r\n\r\nScrisoarea lui Pavel către Biserica din Colose s-a născut în timpuri grele, cînd învăţătura curată a Evangheliei era sub atacul pervertitor al ereziilor filosofiei gnostice. Din închisoarea Romei, Pavel aşterne pe hîrtie rînduri de scriere profundă cu implicaţii care trec dincolo de timp şi de spaţiu prezentîndu-ni-L pe Cristos în copleşitoarea Lui importanţă cosmică. Pentru cititorul de azi, cele scrise de Pavel pot şi trebuie să constituie un ghid înspre înţelegerea persoanei şi lucrării Domnului Isus Cristos.\r\n\r\nPentru titlul, autorul şi data vezi introducerea la Epistola către Efeseni.\r\n\r\nContextul scrierii: Oraşul Colose era situat la aproape 170 de kilometri distanţă de Efes şi îşi înşira locuitorii de-a lungul şerpuitoarei văi a rîului Licus. Pe albia rîului la vale, după un drum de 18 kilometri, ajungeai aproape de vărsarea Licusului în rîul Meander în ţinuturile oraşelor Laodicea şi Ierapole. Aceste două cetăţi stăteau faţă în faţă pe cele două maluri ale rîului şi-şi înlesneau una alteia dezvoltarea.\r\n\r\nLa vremea cînd scrie Pavel, toate cele trei cetăţi erau cuprinse în provincia Romană a Asiei. Locuitorii acestor oraşe erau bogaţi şi afluenţi atît în cele materiale, cît şi în cele culturale. Deşi greu încercaţi de activităţile vulcanice din străfunduri, care le-au pricinuit adeseori cutremure devastatoare, locuitorii n-au vrut să părăsească aceste meleaguri ale belşugului. Mai tîrziu apostolul Ioan avea să mustre Biserica din Laodicea pentru nevegherea ei, pentru superficialitatea credinţei ei şi pentru încrederea ei în lucrurile materiale (Apoc. 3:17-18).\r\n\r\nDespre Biserica din Colose ştim precis că ea nu a fost înfiinţată de Pavel (Col. 2:1). Se pare că cel ce le vestise Evanghelia fusese unul dintre discipolii şi colaboratorii lui Pavel, Epafras (Col. 1:6-7).\r\n\r\nActivitatea acestui lucrător se întindea peste toate Bisericile din valea rîului Licus. Rîvna lui este recunoscută şi lăudată de Pavel (Col. 4:12-13).\r\n\r\nPentru a înţelege mai bine ce se întîmplase în Colose este bine să-i comparăm pe cei de acolo cu exemplul pe care Pavel se străduia să-l dea tuturor. Dacă vă mai aduceţi aminte, Pavel mărturisise celor din Filipi că el „aleargă spre ţintă\", căutînd „să-L cunoască pe Cristos\" şi să se identifice duhovniceşte cu itinerarul parcurs de Cristos în suferinţa, moartea şi învierea Sa. Mesajul apostolului era mărturia unuia care nu se putea opri la jumătate de drum, mulţumindu-se numai cu jumătăţi de măsură. El dorea mereu mai mult şi... mai mult din Cristos. Prin comparaţie cu el, ceea ce se întîmplase în Colose era că cei de acolo nu se puteau mulţumi... numai cu Cristos!\r\n\r\nPavel îl dorea pe Domnul, cei din Colose doreau şi „altceva\" pe lîngă Domnul.\r\n\r\nFără îndoială că în aceste centre ale Asiei pătrunsese învăţătura otrăvită a gnosticilor greci. Despre această filosofic şi despre pretenţiile ei tulburi vom avea ocazia să vorbim mai mult în introducerile la epistolele scrise de apostolul Ioan. Pînă atunci este suficient să spunem că această erezie:\r\n\r\na) ataca persoana şi suficienţa lucrării făcute de Cristos pentru mîntuirea lumii.\r\nb) adăuga Evangheliei învăţături străine care o transformau într-o simplă „alternativă\" a celorlalte „căutări filosofice ale omenirii\".\r\n\r\nGnosticismul era o şcoală filosofică bazată pe o cunoaştere tainică („gnoza\") obţinută printr-un proces de iniţiere în urma căruia li se împărtăşeau discipolilor cunoştinţe despre starea lumii şi despre căile spre refacerea „unităţii şi armoniei cosmice primordiale\". Gnosticii priveau lumea ca pe o succesiune de şapte sfere concentrice ale existenţei cu centrul în Dumnezeu. La mijloc era sfera prezenţei divine, după care se înşirau apoi sferele, din ce în ce mai inferioare, ale „domniilor, dregătoriilor, şi stăpînirilor cereşti\". Născut accidental în cea mai periferică sferă a existenţei, sfera materiei, omul era chemat să caute să se desprindă de ea şi să trăiască în spirit marea aventură a călătoriei prin „sferele cereşti\" înspre „pleroma\" sau locul plinătăţii dumnezeirii.\r\n\r\nFormulată în termeni foarte apropiaţi creştinismului, această învăţătură a gnosticilor a făcut prăpăd în sînul Bisericilor primare. Comentatorii moderni o văd ca pe o încercare tardivă, dar energică a Diavolului de a împiedica răspîndirea mesajului curat al Evangheliei mîntuitoare prin jertfa de la Calvar.\r\n\r\nRedus la cîteva cuvinte, avertismentul lui Pavel spune celor din Colose că nu tot ce seamănă a Evanghelie este Evanghelie. Aduceţi-vă aminte ce le spusese Galatenilor:\r\n\r\n„Mă mir că treceţi aşa de repede de la Cel ce v-a chemat prin harul lui Cristos, la o altă Evanghelie. Nu doar că este o altă Evanghelie; dar sînt unii oameni care vă tulbură şi voiesc să răstoarne Evanghelia lui Cristos. Dar chiar dacă noi înşine sau un înger din cer ar veni să vă propovăduiască o Evanghelie, deosebită de aceea pe care v-am propovăduit-o noi, să fie anatema!\r\n\r\nCum am mai spus, o spun şi acum; dacă vă propovăduieşte cineva o Evanghelie, deosebită de aceea pe care aţi primit-o, să fie anatema!\" (Gal.1:6-9).\r\n\r\nErezia care pătrunsese în Colose era o încrucişare de noţiuni filosofice orientale şi de restricţii alimentare iudaice. Gnosticismul reuşise să strîngă sub acelaşi acoperiş idei contrare şi nu de puţine ori contradictorii. Nu era pentru prima oară şi vai, nici pentru ultima oară, cînd speculaţii rafinate filosofice se uneau cu regulile rigide ale unei religii a faptelor şi ritualelor exterioare. Extremele se ating în eroare pentru că, aşa cum spunea cineva: „dacă mergi suficient de mult înspre est, ajungi cu siguranţă la vest\".\r\n\r\nBiserica trebuie să ştie să se ferească de atacurile învăţăturilor greşite. Iată un pasaj din instrucţiunile pe care i le dă Pavel mai tînărului Timotei:\r\n\r\n„Dar Duhul spune lămurit că în vremurile din urmă, unii se vor lăsa de credinţă, ca să se lipească de duhuri înşelătoare şi de învăţăturile dracilor, abătuţi de făţărnicia unor oameni care vorbesc minciuni, însemnaţi cu fierul roş în însuşi cugetul lor. Ei opresc căsătoria şi întrebuinţarea bucatelor, pe care Dumnezeu le-a făcut ca să fie luate cu mulţumiri de către ceice cred şi cunosc adevărul\" (1 Timotei 4:1-3).\r\n\r\nConţinutul scrierii: Întreaga epistolă către cei din Colose este un răspuns dat atacului învăţătorilor eretici. Mesajul central al scrisorii este: „Cristos, plinătatea Dumnezeirii\", iar intenţia declarată a apostolului este aceea de a-i convinge pe cei ce L-au primit pe Cristos că, dacă îl au pe El, nu mai au nevoie de nimic altceva.\r\n\r\nPunctul central al întregii scrisori se găseşte în versetul 9 din capitolul 2. În acest text, Pavel aşează „pleroma\" de care vorbeau gnosticii în chiar fiinţa lui Cristos:\r\n\r\n„Luaţi seama ca nimeni să nu vă fure cufilosofia, şi cu o amăgire deşartă, după datina oamenilor, după învăţăturile începătoare ale lumii, şi nu după Cristos. Căci în El locuieşte trupeşte toată plinătatea Dumnezeirii. Voi aveţi totul deplin în El, care este capul oricărei domnii şi stăpîniri\" (Col. 2:8-10).\r\n\r\nApostolul prezintă planul lui Dumnezeu de a reface „unitatea şi armonia cosmică primordială\" nu prin ceea ce spuneau gnosticii, ci prin Cristos:\r\n\r\n„El ne-a izbăvit de sub puterea întunerecului şi ne-a strămutat în împărăţia Fiului dragostei Lui, în care avem răscumpărarea, prin sîngele Lui, iertarea păcatelor.\r\n\r\nEl este CHIPUL DUMNEZEULUI CELUI NEVĂZUT, cel întîi născut din toată zidirea. Pentrucă prin El au fost făcute toate lucrurile care sînt sub ceruri şi pe pămînt, cele văzute şi cele nevăzute; fie scaune de domnii, fie dregătorii, fie domnii, fie stăpîniri. Toate au fost făcute prin El şi pentru El.\r\n\r\nEl este mai înainte de toate lucrurile şi toate se ţin prin El.\r\n\r\nEl este Capul trupului, al Bisericii.\r\n\r\nEl este începutul, cel dintîi născut dintre cei morţi, pentru ca în toate lucrurile să aibă întîietatea.\r\n\r\nCăci Dumnezeu a vrut ca toată plinătatea să locuiască în El, şi SĂ ÎMPACE TOTUL CU SINE prin El, atît ce este pe pămînt cît şi ce este în ceruri, făcînd pace, prin sîngele crucii Lui\" (Col. 1: 13-20).\r\n\r\nPentru Pavel, rezolvarea problemei umane, nu este „aventuroasa călătorie prin sferele cereşti\", ci primirea lui Cristos în inimă, căci, întrucît omul nu putea ajunge singur la slavă, Dumnezeu a coborît slava la nivelul nostru şi ne-a oferit-o prin Cristos:\r\n\r\n„Vreau să zic: taina ţinută ascunsă din veşnicii şi în toate veacurile, dar descoperită acum sfinţilor Lui, cărora Dumnezeu a voit să le facă cunoscut care este bogăţia slavei tainei acesteia între neamuri, şi anume: Cristos în voi, nădejdea slavei!\" (Col. 1:26-27).\r\n\r\nConştient de existenţa unor „domnii şi stăpîniri\" care se opun lui Dumnezeu prin învăţături greşite şi care vor să-I zădărnicească planurile, Pavel scrie:\r\n\r\n„A dezbrăcat domniile şi stăpînirile, şi le-a făcut de ocară înaintea lumii, după ce a ieşit biruitor asupra lor prin cruce\" (Col. 2:15).\r\n\r\nÎn ce priveşte „datinile religioase\", Pavel îi absolvă pe creştini de necesitatea respectării lor:\r\n\r\n„Nimeni dar să nu vă judece cu privire la mîncare sau băutură, sau cu privire la o zi de sărbătoare, cu privire la o lună nouă, sau cu privire la o zi de Sabat, care sînt umbra lucrurilor viitoare, dar trupul este al lui Cristos\" (Col. 2:16-23).\r\n\r\nSoluţia tuturor problemelor de viaţă creştină este aşezată de Pavel în viaţa de înviere care trebuie să pulseze în vinele noastre. La sfinţenia creştină se ajunge nu prin „îmbunătăţirea şi disciplinarea firii pămînteşti\", ci prin primirea unei naturi din Cristos care va rodi de la sine:\r\n\r\n„Dacă deci aţi înviat împreună cu Cristos, să umblaţi după lucrurile de sus, unde Cristos şade la dreapta lui Dumnezeu. Gîndiţi-vă la lucrurile de sus, nu la cele de pe pămînt. Căci voi aţi murit şi viaţa voastră este ascunsă cu Cristos în Dumnezeu\" (Col. 3:1-3).\r\n\r\nCUPRINSUL CĂRŢII\r\n\r\nCuvînt de introducere (1:1-8)\r\nRugăciune pentru „umplere\" şi „purtare într-un chip vrednic\" (1:9-14)\r\n\r\n1. DOCTRINĂ - „să vă umpleţi\" (cap.1-2)\r\nCristos - plinătatea lui Dumnezeu în creaţie (1:15-18)\r\nCristos - plinătatea lui Dumnezeu în răscumpărare (1:19-23)\r\nCristos - plinătatea lui Dumnezeu în Biserică (1:24-2:7)\r\nCristos - plinătatea lui Dumnezeu în luptă cu erezia (28-23)\r\n\r\n2. PRACTICA - „să vă purtaţi într-un chip vrednic\" (cap.3-4)\r\nViaţa nouă - şi fiecare credincios (3:1-11)\r\nViaţa nouă - şi credincioşii între ei (3:12-17)\r\nViaţa nouă - şi relaţiile de familie (3:18-21)\r\nViaţa nouă - şi relaţiile de muncă (3:22-4:1)\r\nÎncheiere (4:2-18)\r\nCauta in biblie\r\nTextul cautat\r\nCauta...\r\nCum caut\r\nToate Cuvintele\r\nUnde caut:\r\nToată biblia\r\nVersetul Zilei\r\nJudecatori 6:14-16\r\n\r\nDu-te cu puterea aceasta pe care o ai, si izbaveste pe Israel din mana lui Madian! oare nu te trimit Eu\r\nMeditatia Zilei\r\n\r\nPowered by Biblia Online\r\nFacebook\r\n\r\nParteneri\r\nStiri Crestine\r\nMuzica Crestina\r\nVersuri Crestine\r\nChristian Lyrics\r\nBiblia Online Cornilescu\r\nVersuri Crestine\r\nNu este Crăciun fără Isus - Estera & Laura Bretan\r\nÎn iesle azi - Diana Scridon Pop\r\nBucurie in suflet - Familia Timofte\r\nSe-aude Glas Peste Carpați / Tăria Noastră Fii Mereu - The Beuca Family\r\nUn singur Dumnezeu - Otto Pascal & Biji\r\nFiu iubit - BBSO\r\nClipa de clipa - Philadelphia Band\r\nPrintul Pacii a venit - Alin și Emima Timofte & TB Music Family\r\nPraise - Lumina Worship\r\nUnde ești țara mea?\r\nPowered by VersuriCrestine.ro\r\n \r\n© 2025 Biblia " },
    //1 TESTALONICENI
    { "1 TESALONICENI", "Explicatia Cartii: 1 Tesaloniceni\r\n\r\n„Binecuvîntata noastră nădejde!\"\r\n\r\nCu „1 şi 2 Tesaloniceni\" am ajuns la ultimele dintre epistolele dedicate Bisericii dintre neamuri. De la Romani şi pînă la Tesaloniceni am urmărit dezvoltarea învăţăturilor date nouă de Dumnezeu prin apostolul neamurilor: Pavel. Încununarea învăţăturii lui Pavel este învăţătura despre revenirea Domnului Isus, iar 1 şi 2 Tesaloniceni se ocupă tocmai cu această problemă.\r\n\r\nTitlul: În originalul grec, cartea este numită „Pros Thessalonikeis A\" - „Către Tesaloniceni A\", deoarece ea a fost aşezată pereche cu cea de a doua epistolă a lui Pavel către cei din Tesalonic.\r\n\r\nAutorul: În legătură cu începuturile Bisericii din Tesalonic şi cu implicaţiile lui Pavel în dezvoltarea ei este bine să recitim textul din Fapte 17:1-14. O citire paralelă a pasajelor din Fapte 17:13-16; 18:1-5; şi 1 Tesaloniceni 3:1-8 ne va ajuta şi mai mult să ne formulăm o imagine de ansamblu.\r\n\r\nPe scurt, ceea ce ştim despre Biserica din Tesalonic este faptul că ea a fost născută de Pavel în urma propovăduirii în sinagoga evreiască din Tesalonic. Deşi numărul celor convertiţi dintre iudei nu a fost prea mare, Biserica a ajuns destul de numeroasă prin adăugarea unui număr mare de credincioşi dintre neamuri.\r\n\r\nSuccesul propovăduirii lui Pavel a provocat invidia şi alarmarea comunităţii evreieşti, care s-a pus în mişcare ca un stup de viespi. Ei au agitat întreaga cetate, au tocmit un număr de „oameni fără căpătîi\" şi s-au năpustit asupra creştinilor. Intenţia lor a fost să se răfuiască cu Pavel şi cu Sila, însă...\r\n\r\n„Fiindcă nu i-au găsit, au tîrît pe Iason şi pe cîţiva fraţi înaintea dregătorilor cetăţii şi strigau: „Oamenii aceştia, care au răscolit lumea, au venit şi aici, şi Iason i-a găzduit. Ei toţi lucrează împotriva poruncilor Cezarului, şi spun că este un alt împărat: Isus\" (Fapte 17:-7)\r\n\r\nEste clar că ei au luat din Evanghelia lui Pavel numai ceea ce le-a convenit, transformînd îngrijorarea lor religioasă într-o presupusă problemă politică cu implicaţii sociale. Viclenia lor a reuşit şi astfel asupra creştinilor din Tesalonic a început foarte de timpuriu un val de persecuţie.\r\n\r\nPavel şi Sila s-au deplasat la Berea şi se părea că propovăduirea lor va avea un alt efect asupra comunităţii iudaice, dar „viespile turbate\" din Tesalonic au venit şi acolo:\r\n\r\n„Dar iudeii din Tesalonic, cînd au auzit că Pavel vestea Cuvîntul lui Dumnezeu şi în Berea, au venit acolo, ca să turbure şi să aţîţe noroadele\" (Fapte 17:13).\r\n\r\nAşa că Pavel a trebuit să plece şi din Berea, îndreptîndu-se către Atena.\r\n\r\nPutem spune deci că Biserica din Tesalonic era compusă în majoritate din credincioşi dintre neamuri, cu un mic grup de evrei care crezuseră în cele spuse de Pavel, dar cu întreaga comunitate evreiască din jurul Sinagogii pusă în mişcare într-o aprinsă lucrare de persecuţie.\r\n\r\nData: Pavel scrie această epistolă în preajma anului 51 d.Cr. din oraşul Corint.\r\n\r\nContextul scrierii: În timpul lui Pavel, Tesalonic era un port important şi în acelaşi timp oraşul capitală al provinciei Macedonia. Grecul Casander a întărit această aşezare în anul 315 î.Cr. şi l-a numit după numele soţiei sale, Tesalonica, o soră vitregă a lui Alexandru Macedon. Romanii au cucerit apoi Macedonia în anul 168 î.Cr. şi au reorganizat întreaga provincie stabilind capitala la Tesalonic. În timpul domniei lui Cezar August, oraşul a primit statutul special de „oraş liber\" şi a fost condus de un grup de magistraţi, numiţi şi „politarhi\". Popolaţia a crescut la peste 200.000 de oameni, iar poziţia maritimă i-a înlesnit o bunăstare înfloritoare. Oraşul supravieţuieşte şi astăzi în Grecia, sub numele de „Salonic\".\r\n\r\nCele două scrisori adresate celor din Tesalonic ne prezintă un Pavel care stătuse îndelung de vorbă cu noii convertiţi despre lucrurile viitoare şi mai ales despre revenirea Domnului:\r\n\r\n„Cît despre venirea Domnului nostru Isus Cristos şi strîngerea noastră laolaltă cu El... Nu vă aduceţi aminte cum vă spuneam lucrurile acestea, cînd eram încă la voi? Şi acum ştiţi bine ce-l opreşte ca să nu se descopere... (ştiţi, pentru că v-am spus! - n.n.)\" (2 Tes. 2:5, 6).\r\n\r\nDe fapt, fiecare capitol al celei dintîi epistole se sfîrşeşte cu o referinţă la venirea Domnului (1 Tes. 1:10; 2:19; 3:13; 4:17; 5:23).\r\n\r\nConţinutul cărţii: După ce Pavel a fost silit să plece în grabă din Tesalonic, inima lui a continuat să fie preocupată de starea şi progresul Bisericii de acolo. Ştirile bune pe care i le-a adus Timotei în urma vizitei făcute acolo, l-au făcut pe apostol să scrie această scrisoare. Conţinutul ei este o îmbărbătare a Bisericii, un răspuns dat unei lipse de înţelegere asupra stării celor morţi la venirea Domnului şi un îndemn la o viaţă de sfinţenie în aşteptarea măreţelor evenimente viitoare. Accentul epistolei este pus pe statornicia credinţei în Cristos (1 Tes. 3:8) şi pe o continuă creştere în dragoste şi curăţie (1 Tes. 1:3-10; 2:12-20; 3:10-13; 4:1-5:28). Cartea se ocupă şi cu celelalte aspecte ale vieţuirii creştine, totul fiind impulsionat de această anticipaţie fericită a revenirii Domnului Isus în slavă.\r\n\r\nNimeni nu a spus vreodată că învăţătura despre lucrurile viitoare şi mai ales cea despre venirea Domnului nu va naşte greşeli de interpretare şi de aplicare în practică. Ar fi fost chiar de mirare dacă Diavolul nu s-ar fi năpustit cu înverşunare să anihileze extraordinarul impact al acestor învăţături.\r\n\r\nSatan are două metode de împiedicare a planurilor lui Dumnezeu. Una este este aceea de a-i opri, pur şi simplu, pe oameni să ia cunoştinţă de lucrările lui Dumnezeu, iar a doua constă în pervertirea învăţăturilor primite de oameni din partea lui Dumnezeu.\r\n\r\n\r\nCredincioşii din Tesalonic aflaseră de la Pavel despre lucrurile viitoare şi despre venirea Domnului, aşa că Diavolul nu mai putea împiedica această revelare. Tot ceea ce-i mai rămăsese de făcut era să încerce să denatureze această învăţătură şi să-i facă pe cei credincioşi să o aplice în mod greşit în practică.\r\n\r\nProblema celor din Tesalonic s-a ivit atunci cînd unii dintre credincioşii adunării au murit şi Diavolul a semănat în inimile celor rămaşi deznădejdea şi tristeţea. Cumva, cei din Tesalonic ştiau că Domnul îi va lua la cer, dar nu le era deloc clară soarta celor care vor muri pînă la venirea Domnului. Pavel le răspunde la această nedumerire în 1 Tes. 4:13-18.\r\n\r\nEste evident că greşelile de interpretare nu se elimină prin restrîngerea învăţăturii, ci prin mai multă învăţătură. Aceasta a fost metoda folosită de Pavel. El nu a spus un fel de: „Iertaţi-mă fraţilor! Îmi dau seama că v-am băgat în lucruri pe care nu este bine să le ştiţi. Mai bine nu vă spuneam despre venirea Domnului\", ci i-a condus pe cei din Tesalonic la o înţelegere mai profundă.\r\n\r\nRemarcaţi autoritatea celor spuse de Pavel:\r\n\r\n„Iată, în adevăr, ce vă spunem prin Cuvîntul Domnului...” (1 Tes. 4:15).\r\n\r\nAici nu mai este vorba despre păreri personale sau despre interpretări de ale lui Pavel. Domnul Isus personal i-a încredinţat acest mesaj pentru Biserică!\r\n\r\nCuvinte cheie şi terne caracteristice: Epistola cuprinde descoperirea unor aspecte ale „răpirii\" Bisericii la venirea Domnului, pe care nu le mai găsim în nici o altă epistolă. Textul ne dă o cronologie a secvenţelor în care se vor petrece evenimentele legate de învierea celor morţi, transformarea celor vii şi ridicarea credincioşilor „în nori\" pentru a-L întîlni acolo pe Domnul (1 Tes. 4:16-18).\r\n\r\nCUPRINSUL CĂRŢII\r\n1 Tesaloniceni - Cristos, nădejdea noastră.\r\n\r\nCuvînt de salut 1:1\r\n\r\nI. PRIVIND ÎNAPOI: Cum au fost ei salvaţi (cap. 1-3)\r\n\r\na) CONVERTIRE EXEMPLARĂ (cap.1)\r\nAu cunoscut puterea Evangheliei (5), au devenit exemple (6, 7) şi martori (8-10)\r\n\r\nb) EVANGHELISM EXEMPLAR (cap.2)\r\nÎn motivare (1-6), în conduită (7-12), în mesaj (13-16)\r\n\r\nc) PURTARE DE GRIJĂ EXEMPLARĂ (cap.3)\r\nPreocupare (1-5), creştere (6-8), rugăciune fierbinte (9-13)\r\n\r\nII. PRIVIND ÎNAINTE: Cum trebuie ei să trăiască (cap.4-5)\r\n\r\na) Chemare şi conduită (4:1-12)\r\nÎn lumina voiei lui Dumnezeu.\r\n\r\nb) Mîngîiere şi îndemn (5:12-24)\r\nÎn aşteptarea venirii Domnului.\r\n\r\nc) Rînduială şi credincioşie (5:12-24)\r\nÎn trăirea părtăşiei creştine.\r\n\r\nCerere şi benedicţia (5:25-28)" },
    //2 TESTALONICENI
    { "2 TESALONICENI", "Explicatia Cartii: 2 Tesaloniceni\r\n\r\nAşa cum am arătat deja, cele nouă „Epistole ale Bisericii Creştine\" sînt aşezate într-o succesiune de trei grupe de 4 epistole, 3 epistole şi respectiv 2 epistole. Primele patru vorbesc despre CRUCE ca esenţă a doctrinei despre mîntuire, următoarele trei vorbesc despre BISERICĂ ca trup spiritual armonios, compus din cei care au crezut în Cristos, iar ultimele două, aşezate pereche, prezintă VENIREA DOMNULUI ca eveniment sigur şi sursă de motivaţie pentru perseverenţa sfinţilor. În primele patru, credinţa priveşte înapoi la Cruce şi este întărită. În cele trei de la mijloc, dragostea priveşte sus la Mirele ceresc şi creste în devotament. În ultimele două, cele scrise tesalonicenilor, nădejdea priveşte înainte spre sfîrşitul care se apropie şi se aprinde de dor.\r\n\r\nTitlul: În originalul grec, cartea poartă numele de „Pros Tessalonikeis B\" - „Către tesaloniceni B\".\r\n\r\nAutorul: Pavel scrie această epistolă ca o revenire asupra unor probleme care fuseseră deja discutate (2 Tes. 2:5), dar fuseseră între timp denaturate de interpretări mincinoase.\r\n\r\nData: Probabil tot în jurul anului 51 d.Cr. la un anumit interval de timp după scrierea primei epistole.\r\n\r\nContextul scrierii: Pentru circumstanţele istorice, recitiţi ceea ce a fost scris în introducerea făcută primei epistole. Din punct de vedere spiritual, această a doua scrisoare s-a născut datorită unui „fals\" viclean prin care „cineva\" alterase învăţătura despre ziua revenirii Domnului.\r\n\r\nConţinutul cărţii: Cea de-a doua epistolă către Tesaloniceni, este o urmare firească a celei dintîi, în care Pavel le prezentase credincioşilor adevărul despre venirea zilei Domnului (1 Tes. 5:1-11). La scurtă vreme după citirea acelei epistole, în Biserica din Tesalonic se întîmplase însă ceva neprevăzut. Profitînd de faptul că Pavel nu-şi scria el însuşi corespondenţa din cauza bolii lui de ochi, ci o dicta altora (Rom. 16:22; 1 Cor. 16:21), „cineva\" s-a găsit să scrie o „epistolă\" pastorală plină de erezii despre venirea zilei Domnului. Din cauză că nu-i cunoşteau scrisul de mînă, pentru o vreme falsul a trecut neobservat. Cînd Pavel a aflat, s-a grăbit să le scrie cea de a doua epistolă. Aşa a apărut 2 Tesaloniceni.\r\n\r\n„Cît priveşte venirea Domnului nostru Isus Cristos şi strîngerea noastră laolaltă cu El, vă rugăm, fraţilor, să nu vă lăsaţi clătinaţi aşa de repede în mintea voastră, şi să nu vă tulburaţi de vreun duh, nici de vreo vorbă, nici de vreo epistolă, ca venind de la noi, ca şi cum ziua Domnului ar fi venit chiar. Nimeni să nu vă amăgească în vreun chip...” (2 Tesal. 2:1-3).\r\n\r\nPentru ca lucrarea de rătăcire să nu se mai repete, Pavel vrea să-i înarmeze pe tesaloniceni cu un semn de recunoaştere şi de verificare a epistolelor sale:\r\n\r\n„Urarea de sănătate este scrisă cu mîna mea: Pavel. Acesta este semnul în fiecare epistolă; aşa scriu eu\" (3:17).\r\n\r\nErezia adusă de scrisoarea „apocrifă\" era aceea că „ziua Domnului ar fi şi venit chiar\".\r\n\r\nSensul era că Domnul nu se va întoarce vizibil şi trupeşte, ci ar fi vorba despre o reîntoarcere „în Duhul\", ca ceea petrecută la Rusalii, în această interpretare, „ziua Domnului\" nu ar mai trebui aşteptată ca un eveniment glorios aşezat în viitor, ci ea ar trebui înţeleasă ca o vreme de har, o perioadă de „o zi\" în calendarul lui Dumnezeu în care „o zi este ca o mie de ani şi o mie de ani sînt ca o zi\". Din această perspectivă, ziua Domnului era interpretată ca fiind vremea Bisericii.\r\n\r\nDacă veţi asculta cu atenţie în jur, eroarea aceasta mai persistă şi în zilele noastre.\r\n\r\nCorecţia făcută de Pavel este energică şi imediată. El spune că „ziua Domnului\", ca eveniment unic în planul lui Dumnezeu, nu va veni decît după ce pămîntul va cunoaşte două evenimente catrastrofice:\r\n\r\na. „lepădarea de credinţă\" (2:3), înţeleasă ca un refuz mondial al ofertelor lui Dumnezeu de rezolvare a problemelor lumii, ca o totală apostazie în care umanitatea îi va întoarce spatele lui Dumnezeu şi\r\n\r\nb. „descoperirea omului fărădelegii\" numit şi, „fiul pierzării, potrivnicul, care se înalţă mai presus de tot ce se numeşte „Dumnezeu\", sau de ce este vrednic de închinare. Aşa că se va aşeza în Templul lui Dumnezeu, dîndu-se drept Dumnezeu\" (2:3).\r\n\r\nÎntr-un sens foarte vag, omenirea trăia încă din vremea lui Pavel evenimente pregătitoare acestei apariţii mondiale:\r\n\r\n„Căci taina fărădelegii a şi început să lucreze; trebuie numai ca cel ce o opreşte acum, să fie luat din drumul ei\" (2 Tes. 2:7)\r\n\r\nTotuşi, nimic din ceea ce s-a petrecut atunci sau din ceea ce se petrece acum nu se poate compara cu ce va fi atunci cînd: „...se va arăta acel Nelegiut...” (2 Tes. 2:8)\r\n\r\n„Arătarea lui se va face prin puterea Satanei, cu tot felul de minuni, de semne şi de puteri mincinoase, şi cu toate amăgirile nelegiuirii pentru cei cesîntpe calea pierzării, pentru că n-au primit dragostea adevărului ca să fie mîntuiţi\" (2:9-10).\r\n\r\nLucrarea acestui „Anticrist\" va fi îngăduită de Dumnezeu ca o pedeapsă trimeasă asupra lumii care L-a refuzat pe Cristos. Dumnezeu le va da ceea ce au cerut de-a lungul secolelor, o lume fără Cuvîntul lui Dumnezeu, fără opreliştile prezenţei Duhului lui Dumnezeu; o lume cufundată în părtăşia celui care însumează toate realizările „separării şi împotrivirii faţă de Dumnezeu\": Satan însuşi.\r\n\r\nDomnia lui Anticrist va fi pregătirea decorului pentru scena finală a istoriei omenirii, cînd Dumnezeu va aduce pedeapsa divină asupra lumii păcătoase. Lui Satan îi va fi îngăduit să păşească în arenă pe faţă, identificîndu-se cu lumea înşelată de el, conducînd-o şi însufleţind-o într-o ultimă zvîrcolire împotriva dumnezeirii.\r\n\r\n„Din această pricină, Dumnezeu le trimite o lucrare de rătăcire, ca să creadă o minciună; pentru ca toţi cei ce n-au crezut adevărul, ci au găsit plăcerea în nelegiuire, să fie osîndiţi\" (2 Tes. 2:11-12)\r\n\r\nDumnezeu nu va lăsa nimănui sarcina de a-L confrunta pe „potrivnicul\" Său. Înfruntarea finală va fi între Diavol şi Domnul Isus însuşi, iar victoria Domnului va fi deplină: „pe care Domnul Isus îl va nimici cu suflarea gurii Sale, şi-l va prăpădi cu arătarea venirii Sale\" (2 Tes. 2:8).\r\n\r\nO altă greşală a celor din Tesalonica a fost „trăirea în neorînduială\". Bazaţi pe faptul că venirea Domnului este aproape, cei ce aveau înclinaţii spre „lene\" au găsit pretext pentru părăsirea ocupaţiilor zilnice şi pentru începerea unui trai „de pe o zi pe alta\". Astfel de oameni deveniseră o povară pentru adunare şi o proastă mărturie faţă de cei de afară. Nu este de mirare că Pavel i-a mustrat cu asprime: „Cine nu vrea să muncească, nici să nu mănînce\" (2 Tes. 3:6-15). Venirea Domnului nu este o scuză pentru leneşi. Creştinul trebuie să-şi slujească Domnul pînă în cea din urmă clipă a existenţei sale.\r\n\r\nCuvinte cheie şi teme caracteristice: Una dintre cele mai interesante remarci ale lui Pavel este aceea din 2 Tes. 2:7: „Căci taina fărădelegii a şi început să lucreze; trebuie numai ca cel ce o opreşte acum, să fie luat din drumul ei\".\r\n\r\nPe cine prezintă apostolul cu aceste cuvinte? Cea mai probabilă interpretare este aceea că Pavel a vorbit aici despre prezenţa şi lucrarea Duhului Sfînt în dispensaţia Bisericii. Se cuvine să-i mulţumim împreună lui Dumnezeu pentru că în vremea de acum a pus o „piedică\" în calea Diavolului. Slavă Domnului că Satan nu poate face tot ceea ce voieşte!\r\n\r\nÎn timpul Mileniului, Satan va fi legat „completamente\", scos din activitate şi ţinut pentru o vreme în „abis\", dar, pînă atunci, la sfîrşitul perioadei de har în care ne aflăm „cel ce o opreşte acum\" îi va fi luat din cale. Libertatea de lucrare a Diavolului va creşte şi va fi: „Vai de voi, pămînt şi mare! Căci Diavolul s-a pogorît la voi, cuprins de o mînie mare, fiindcă ştie că are puţină vreme\" (Apoc.l2:12).\r\n\r\nAbia atunci îşi va da el pe faţă întreaga răutate şi putere de distrugere, căci nu va mai fi limitat să caute să-i înşele pe sfinţi prefăcîndu-se într-un „înger de lumină\" (2 Cor.11:14), ci se va putea manifesta în toată cruzimea lui: „ca un leu care răcneşte şi caută pe cine să înghită\" (1 Petru 5:8).\r\n\r\nAtunci, el va apare ca: „fiara ieşită din mare\" şi ca „fiara ieşită din pămînt\" cu chipul omului identificat sub taina numărului 666 (Apoc. 13). Întrupat în această fiinţă umană împuternicită cu resurse nemaivăzute în istoria lumii, Diavolul va uimi întîi lumea înşelînd-o, pentru ca apoi să o stăpînească chinuind-o.\r\n\r\nDe ce „trebuiesc\" să se întîmple toate acestea? Pentru că numai în felul acesta omenirea, care L-a refuzat pe Cristos, va învăţa pe propria ei piele ce groaznică este vieţuirea fără prezenţa Domnului, într-o ultimă şi culminantă lecţie, Dumnezeu va arăta tuturor ce înseamnă să te răscoli „împotriva Domnului şi împotriva Unsului Său\" (Ps. 2 lămurit în Fapte 4:25-28).\r\n\r\nDar chiar atunci cînd se va părea că 666 a învins toată creaţia, se va arăta din cer desăvârşitul „7” al Dumnezeirii. Mîntuitorul lumii se va coborî în slavă, vizibil „ca fulgerul care răsare de la răsărit şi se vede pînă la apus\", şi-l va nimici pe Diavol „cu suflarea gurii Sale şi cu arătarea venirii Sale\" (2 Tesal. 2:8).\r\n\r\nCUPRINSUL CĂRŢII\r\n\r\n„Aşteptînd, veghind şi lucrînd cu răbdare\"\r\n\r\nCuvînt de salut (1:1, 2)\r\n\r\nI. ALINARE - în nădejdea revenirii Lui (cap. 1)\r\nOdihni într-o viaţă de încercări (3-7)\r\nRăsplăti în viitor la venirea Lui (8-12)\r\n\r\nII. AVERTIZARE - în privinţa timpului venirii Lui (cap. 2)\r\nCînd şi cum va veni (1-12)\r\nDe ce şi cum să aşteptăm (13-17)\r\n\r\nIII. ANGRENARE - în pregătire pentru venirea Lui (cap. 3)\r\nÎncurajare pentru ce este bun (1-5)\r\nMustrare pentru ce este rău (6-15)\r\nBenedicţie şi semnătură (3:16-18)\r\nCauta in biblie\r\nTextul cautat\r\nCauta...\r\nCum caut\r\nToate Cuvintele\r\nUnde caut:\r\nToată biblia\r\nVersetul Zilei\r\nJudecatori 6:14-16\r\n\r\nDu-te cu puterea aceasta pe care o ai, si izbaveste pe Israel din mana lui Madian! oare nu te trimit Eu\r\nMeditatia Zilei\r\n\r\nPowered by Biblia Online\r\nFacebook\r\n\r\nParteneri\r\nStiri Crestine\r\nMuzica Crestina\r\nVersuri Crestine\r\nChristian Lyrics\r\nBiblia Online Cornilescu\r\nVersuri Crestine\r\nNu este Crăciun fără Isus - Estera & Laura Bretan\r\nÎn iesle azi - Diana Scridon Pop\r\nBucurie in suflet - Familia Timofte\r\nSe-aude Glas Peste Carpați / Tăria Noastră Fii Mereu - The Beuca Family\r\nUn singur Dumnezeu - Otto Pascal & Biji\r\nFiu iubit - BBSO\r\nClipa de clipa - Philadelphia Band\r\nPrintul Pacii a venit - Alin și Emima Timofte & TB Music Family\r\nPraise - Lumina Worship\r\nUnde ești țara mea?\r\nPowered by VersuriCrestine.ro\r\n \r\n© 2025 " },
    // 1 TMOTEI
    { "1 TIMOTEI", "Explicatia Cartii: 1 Timotei\r\n\r\nTitlul: În originalul grec, cartea poartă numele: „Pros Timotheon A\" - „Către Timotei - A\". Despre Timotei ştim că a fost un copil născut dintr-o căsătorie „mixtă\" între un tată grec şi o mamă evreică (Fapte 16:1). Încă din cea mai fragedă pruncie, Timotei a luat cunoştinţă cu scrierile sfinte evreieşti prin mama lui, numită Eunice şi prin bunica lui, numită Lois (1 Tim. 1:5; 3:15). S-a convertit cu ocazia primei vizite a lui Pavel la Listra (1 Cor. 4:17; 1 Tim. 1:2; 2 Tim. 1:2). Cu ocazia celei de a doua vizite în Listra, Pavel hotărăşte să-l ia cu el şi-l taie împrejur din pricina iudeilor (Fapte 16:1-3). Este clar că pînă la ceasul acela, Timotei fusese privit de evrei ca un produs „compromis\" al unei căsătorii „nelegiuite\" şi crescut de un tată care n-a vrut să respecte asezămintele tradiţionale ale religiei evreilor. Crescut de Pavel şi ordinat în lucrarea creştină (1 Tim. 4:14; 2 Tim. 1:6), Timotei devine colaboratorul şi însoţitorul apostolului în călătoriile sale prin Troa, Berea, Tesalonic şi Corint (Fapte 16 - 18; 1 Tes. 3:1, 2). În timpul celei de a treia călătorii misionare, Timotei lucrează împreună cu Pavel sau este trimis ca reprezentant al apostolului în Efes, Macedonia şi Corint. Timotei a stat cu Pavel în timpul primei lui detenţii la Roma şi s-a înapoiat împreună cu apostolul la Filipi (Filip. 2:19-23). Mai tîrziu, Pavel îl lasă la Efes pentru a supraveghea lucrarea de acolo (1 Tim. 1:3), dar cînd este iarăşi închis, apostolul îl roagă să vină alături de el la Roma (2 Tim. 4:9, 21). Textul din Evrei 13:23 ne spune că Timotei însuşi a călcat pe urmele apostolului gustînd din viaţa amară a închisorii.\r\n\r\nPersonalitatea lui Timotei este remarcabilă. Bolnăvicios (1 Tim. 5:23), timid (1 Tim. 1:7) şi fără experienţă (1 Tim. 4:12), el a lucrat cu devotament alături de Pavel, dovedind abilitate de predicator, credincioşie de ucenic, rîvnă de apostol şi perseverenţă de misionar.\r\n\r\nAutorul: Epistolele pastorale sînt o expresie a grijei apostolului Pavel pentru Biserici şi pentru ucenicii pe care el i-a promovat în lucrarea Evangheliei. „Tu dar, copilul meu, întăreşte-te în harul care este în Cristos Isus. Şi ce ai auzit de la mine în faţa multor marturi, încredinţează la oameni de încredere, care să fie în stare să înveţe şi pe alţii\" (2 Tim. 2: 1-2). Extraordinara răspîndire a creştinismului din primele veacuri poate fi explicată numai dacă ţinem seama că atunci Biserica se răspîndea nu prin „teologie\", ca astăzi, ci prin „ucenicie\" (Fapte 18:23-28).\r\n\r\nData: Pavel scrie această scrisoare în preajma anului 63 d.Cr. la puţin timp după eliberarea lui din închisoarea de la Roma. Trecuseră 7 ani de cînd apostolul îi avertizase pe presbiterii bisericii din Efes de pericolul „lupilor răpitori\" care se vor strecura în Biserică şi nu vor cruţa turma (Fapte 20:29, 30). Temerile apostolului se confirmaseră între timp. Urmaşii lui Imeneu şi Alexandru, pe care Pavel îi dăduse pe mîna Satanei, ca să se înveţe să nu hulească (1 Tim 1:20-21) ridicaseră din nou capul (1 Tim. 4:1-3). Acum în Efes era lucrător Timotei şi Pavel revine în scrisoarea pe care o trimite asupra aceloraşi teme discutate cu episcopii Bisericii.\r\n\r\nContextul scrierii: Timotei primeşte aceste epistole ale lui Pavel în timpul şederii lui în Biserica din Efes. Pavel lucrase cu mult spor în oraşul acela şi inima lui era legată tare de credincioşii de acolo. De fapt, lucrarea lui Dumnezeu în Efes a fost aşa de puternică, iar numărul convertiţilor la creştinism a fost aşa de mare încît într-un interval de timp de 50 de ani templele păgîne au rămas goale şi multe dintre ele au trebuit închise. Iată ce găsim scris în cronica din cartea Faptele Apostolilor:\r\n\r\n„În urmă, Pavel a intrat în sinagogă, unde vorbea cu îndrăzneală. Timp de trei luni a vorbit cu ei despre lucruri privitoare la Împărăţia lui Dumnezeu şi căuta să înduplece pe cei ce-l ascultau. Dar, fiindcă unii rămîneau împietriţi şi necredincioşi, şi vorbeau de rău Calea Domnului înaintea norodului, Pavel a plecat de la ei, a despărţit pe ucenici de ei, şi a învăţat în fiecare zi pe norod în şcoala unuia numit Tiran. Lucrul acesta a ţinut doi ani, aşa că toţi ceice locuiau în Asia, iudei şi Greci, au auzit Cuvîntul Domnului\" (Fapte 19:8-10).\r\n\r\nFără îndoială că expresia: „toţi ceice locuiau în Asia au auzit Cuvîntul Domnului\" nu înseamnă că toţi locuitorii Asiei au venit ei înşişi la Pavel în Efes sau că toţi s-au înscris în şcoala lui Tiran. Răspîndirea deosebită a Evangheliei s-a făcut prin cei pregătiţi de Pavel în acea şcoală a lui Tiran şi trimişi apoi să propovăduiască mai departe. Tactica lucrării lui Pavel a fost să atingă personal centrele majore din Imperiu şi să instruiască acolo pe lucrătorii care să răspîndească apoi în jur Evanghelia. În vremea aceea nu existau Seminarii şi nici şcoli teologice. Metoda folosită de apostol pentru multiplicarea numărului de vestitori ai Evangheliei este descrisă în cea de a doua epistolă către Timotei:\r\n\r\n„Dar tu, copilul meu, întăreşte-te în harul care este în Cristos Isus. Şi ce ai auzit de la mine, în faţa multor marturi, încredinţează la oameni de încredere, care să fie în stare să înveţe şi pe alţii\" (2 Tim. 2:1-2).\r\n\r\nApostolul îl sfătuieşte pe Timotei să preia tactica lui de lucru şi să „se multiplice\" pe sine, instruind lucrători care să poarte mai departe Evanghelia. În textul de mai sus sînt cuprise patru nivele de lucrători implicaţi în lucrarea creştină:\r\n\r\n1) Pavel - l-a instruit pe Timotei şi pe mulţi alţii („în faţa multor marturi\").\r\n\r\n2) Timotei - este îndemnat să încredinţeze învăţătura primită la alţi „oameni de încredere\".\r\n\r\n3) Aceşti „oameni de încredere\" - vor trebui aleşi după capacitatea for de a fi „în stare să înveţe şi pe alţii\".\r\n\r\n4) Acei „alţii\" necunoscuţi încă, dar care se vor integra apoi în această „ştafetă nevăzută\" a Evangheliei.\r\n\r\nUn alt lucru pe care trebuie să-l spunem despre caracteristicile acelei perioade este faptul că în vremea aceea Bisericile nu aveau clădiri afectate ţinerii de servicii divine. Grupurile de creştini se întruneau în casele celor credincioşi (Rom. 16:5, 23), în aer liber sau în săli luate cu chirie (Fapte 8-10). Clădirile de Biserici au apărut numai după 200 de ani de la moartea lui Pavel, cînd, în urma decretului dat de Constantin cel Mare, a fost încetată persecuţia asupra credincioşilor şi creştinismul a devenit religie de Stat. În vremea lui Timotei, existau sute de „Biserici\" mici care se adunau prin case (Filimon 2), sub călăuzirea unor lideri locali numiţi fie „presbiteri\", fie „păstori\", fie „episcopi\" (Fapte 20:17, 28).\r\n\r\nConţinutul cărţii: Epistolele către Timotei şi Tit sînt veritabile cursuri de „teologie pastorală\". Oricine vrea să-l aibă pe Pavel drept profesor, poate citi aceste lucrări ale lui. În vremea aceea, Timotei a funcţionat ca reprezentant apostolic în Efes şi probabil şi în alte părţi ale Asiei. Misiunea lui a fost să „aşeze presbiteri\", să corecteze învăţăturile greşite şi să supravegheze viaţa Bisericilor înfiinţate de Pavel. Din textul cărţii reiese clar că epistola este o continuare a învăţăturilor pe care Pavel i le-a dat lui Timotei prin viu grai (1 Tim. 1:3-4). Scopul urmărit de Pavel cu colaboratorii săi mai tineri este exprimat foarte clar în 1 Tim. 3:14-15: „Îţi scriu aceste lucruri cu nădejdea că voi veni în curînd la tine. Dar dacă voi zăbovi, să ştii cum trebuie să te porţi în casa Dumnezeului celui viu, stîlpul şi temelia adevărului\".\r\n\r\nCine citeşte cu atenţie scrisorile pastorale scrise de Pavel observă foarte repede că toate sînt un fel de testament al apostolului. Ucenicii lui sînt îndemnaţi în mod repetat să „păstreze\" ceea ce le fusese încredinţat (1 Tim. 1:18-19; 3:9; 6:14, 20; 2 Tim. 1:13, 14; 2:2). Averea lăsată de apostol urmaşilor săi este identificată în 1 Tim. 1:11: „Evanghelia slavei fericitului Dumnezeu care mi-a fost încredinţată mie\". Epistolele pastorale sînt o ilustrare a schimbului de ştafetă dintre două generaţii de lucrători: Pavel i-a crescut pe colaboratorii săi prin exemplu personal şi prin învăţătură (Filip. 3:17; 4:9). Acum este rîndul lui Timotei să ducă ştafeta mai departe. Metoda de creştere trebuie să rămînă mereu aceiaşi: exemplul personal şi învăţătura (1 Tim. 4:12-13, 16).\r\n\r\nConţinutul epistolei este foarte clar şi foarte sistematic aşezat: după o scurtă explicare a motivului pentru care a fost scrisă cartea (1 Tim. 1:1-17), urmează îndemnul lui Pavel pentru „păstrarea\" moştenirii spirituale lăsate de Pavel (1 Tim. 1:18-20). Acest „depozit de învăţătură\" este apoi descris în două secţiuni caracteristice: prima parte cuprinde învăţătura despre Biserică şi despre organizarea ei (1 Tim. 2 şi 3), iar a doua parte cuprinde învăţătura despre lucrătorul creştin şi despre felul lui de comportament faţă de diferite categorii de credincioşi din Biserică (1 Tim. 4-6).\r\n\r\nCuvinte cheie şi teme caracteristice: Tema întregii cărţi este sintetizată în următoarea expresie: „Ca să şti cum să te porţi în Biserică\" (1 Tim. 3:15).\r\n\r\n1 Timotei şi Tit sînt cele două epistole care ne prezintă caracterul şi caracteristicile liderilor spirituali ai Bisericii: ce sînt ei în ei înşişi (1 Tim. 3:2), ce sînt ei în relaţiile cu alţii (1 Tim. 3:3), ce sînt ei în familiile lor (1 Tim. 3:4-5) şi ce sînt ei în relaţiile cu lumea din jur (1 Tim. 3:6-7).\r\n\r\nCUPRINSUL CĂRŢII „Manualul presbiterilor\"\r\nIntroducere 1:1-17\r\nÎndemn, 1:18-20\r\n\r\nI. Biserica şi organizarea ei 2-3\r\na. Lucrarea în Biserică, 2:1-7\r\nb. Rugăciunea în Biserică, 2:8\r\nc. Poziţia femeii în Biserică, 2:9-15\r\nd. Lucrătorii Bisericii\r\nPresbiterii, 3:1-7\r\nDiaconii, 3:8-14\r\n\r\nII. Conduita lucrătorul creştin 4-6\r\na. În combaterea învăţătorilor mincinoşi, 4:1-6\r\nb. În practicarea evlaviei, 4:7-11\r\nc. Într-o pildă de viaţă şi învăţătură, 4:11-16\r\nd. În relaţiile cu cei tineri şi cu cei bătrîni, 5:1 -2\r\ne. În îngrijirea văduvelor din Biserică, 4:3-16\r\nf. În relaţiile cu ceilalţi presbiteri, 5:17-25\r\ng. În relaţiile cu cei aflaţi în robie, 6:1-8\r\nh. În relaţiile cu cei bogaţi, 6:9-19\r\n\r\nRepetarea îndemnului, 6:20-21" },
    // 2 TIMOTEI
    { "2 TIMOTEI", "Explicatia Cartii: 2 Timotei\r\n\r\nAceastă epistolă este „cîntecul de lebădă\" al lui Pavel.\r\n\r\nTitlul: În originalul grec, cartea poartă numele: „Pros Timotheon B\" - „Către Timotei B\". Pentru descrirea raportului dintre Pavel şi Timotei vă rugăm să citiţi introducerea la „1 Timotei\".\r\n\r\nAutorul: Cel ce scrie aceste rînduri duioase este Pavel, „tatăl spiritual\" al lui Timotei şi al atîtor altora. Bătrîn, obosit, bolnav şi aflat aproape de clipa plecării lui la Domnul, apostolul se mai apleacă încă o dată asupra foii şi scrie cu „lacrimi de suflet\" o epistolă de dragoste creştină.\r\n\r\nData: După eliberarea lui Pavel din prima lui detenţie în închisoarea din Roma (Fapte 28:30), apostolul călătoreşte prin Efes (1 Tim 1:3), Creta (Tit 1:5), Nicopoli (Tit 3:12), Milet (2 Tim 4:20) şi Troa (2 Tim. 4:13). Pavel avea însă de împlinit o profeţie (Fapte 21:11-13), aşa că drumurile vieţii lui îl duc iarăşi în închisoarea Romei (2 Tim. l:16-17), unde îşi aşteaptă judecata şi execuţia (2 Tim. 4:6-8). Datele acestea fixează timpul scrierii celei de a doua epistole către Timotei cam prin anul 66 d.Cr. spre sfîrşitul domniei împăratului roman Nero.\r\n\r\nContextul scrierii: Închisoarea este ultimul loc din care ne-am putea aştepta să primim o scrisoare de încurajare, iar condamnatul la moarte este ultimul om din lume de la care te aştepţi să auzi cuvinte de îmbărbătare. Totuşi, situaţiile acestea paradoxale s-au petrecut întocmai cu Timotei şi cu Pavel.\r\n\r\nConţinutul scrierii: Această ultimă epistolă scrisă de Pavel este un mesaj de întărire şi încurajare din partea apostolului pentru mai tînărul şi mai timidul său colaborator aflat în continuare la Efes. Departe de a se considera un învins, Pavel, ca un veritabil soldat al crucii, îi mai face ultima instrucţie lui Timotei. Veteranul predă armele recrutului înainte de lăsarea sa la vatră (2 Tim. 4:6, 18).\r\n\r\nCuvinte cheie şi teme caracteristice: Tema întregii cărţi poate fi luată din expresia: „un bun ostaş al lui Cristos\" (2 Tim. 2:3). Alte teme importante din epistolă sînt: „Inspirarea Scripturilor\" (2 Tim. 3:16-17), credincioşie pînă la capăt (2 Tim. 8-9), „cununa neprihănirii\" (2 Tim. 4:6-8), mîntuirea finală (2 Tim. 4:18). Nicăieri în altă carte a Noului Testament nu găsim o descriere mai amănunţită a stării caracteristice a omenirii la vremea sfîrşitului (2 Tim. 3:1-9; 4:3-4).\r\n\r\nCUPRINSUL CĂRŢII\r\nChemare la credincioşie\r\n\r\nIntroducere, 1:1-2\r\n\r\nI. Adevăratul lucrător şi încercările prezente 1-2\r\n\r\nChemări personale:\r\na. „înflăcărează\", 1:6\r\nb. „Să nu-ţi fie ruşine de mărturisirea Domnului\", 1:8\r\nc. „întăreşte-te\", 2:1\r\nd. „sufere\", 2:3\r\ne. „adu-ţi aminte de Domnul\", 2:8\r\nf. „adu-ţi aminte de exemplul meu\", 1:12; 2:9-10; 4:5-8\r\n\r\nChemări pastorale:\r\na. „adu-le aminte\", 2:14\r\nb. „roagă-i fierbinte\", 2:14\r\nc. „împarte drept Cuvîntul adevărului\", 2:15\r\nd. „fereşte-te de vorbăriile goale\", 2:16\r\ne. „fereşte-te de întrebările nebune\", 2:23\r\n\r\nII. Adevăratul lucrător şi încercările viitoare 3-4\r\n\r\nChemări personale:\r\na. „fereşte-te de relele lumii\", 3:1-13\r\nb. „rămîi în Cuvîntul Domnului\", 3:14-15\r\nc.„adu-ţi aminte de pilda mea\", 3:10-11\r\nd. „adu-ţi aminte de Scriptură\", 3:16-17\r\n\r\nChemări pastorale:\r\na. „propovăduieşte Cuvîntul\", 4:2\r\nb. „rabdă suferinţele\", 4:5\r\nc. „fă lucrul unui evanghelist\", 4:5\r\nd. „împlineşte-ţi bine slujba\", 4:5\r\ne. „adu-ţi aminte de judecata şi de împărăţia viitoare\", 4:1\r\nf. „adu-ţi aminte de cununa răsplătirii\", 4:8\r\n\r\nÎncheiere, 4:9-22" },
    //TIT
    { "TIT", "Explicatia Cartii: Tit\r\n\r\nTitlul: În originalul grec, cartea poartă numele: „Pros Titon\" - „Către Tit\".\r\n\r\nDestinatarul acestei epistole este unul dintre convertiţii apostolului Pavel (Tit 1:4) şi în acelaşi timp unul dintre colaboratorii săi mai tineri. Importanţa lui Tit în dezvoltarea Bisericii creştine din primul veac este adesea trecută cu vederea. Totuşi, omul acesta a fost unul dintre pilonii lucrării misionare între neamuri. Lucrarea lui misionară s-a întins pînă în Dalmaţia, Yugoslavia de astăzi. Tit este menţionat de 13 ori în cărţile Noului Testament. El a fost primul convertit dintre neamuri, pe care Pavel l-a dat exemplu în adunarea Consiliului Bisericii din Ierusalim: „Nici chiar Tit, care era cu mine, măcar că era grec, n-a fost silit să se taie împrejur, din pricina fraţilor mincinoşi, furişaţi şi strecuraţi printre noi, ca să pîndească slobozenia, pa care o avem în Cristos Isus\" (Gal. 2:1-4).\r\n\r\nDupă ce a fost încercat în alte misiuni în care şi-a dovedit rîvna, maturitatea şi credincioşia, Tit a primit din partea lui Pavel una dintre cele mai importante şi dificile însărcinări: să rămînă în Creta ca să pună „în rînduială ce mai rămîne de rînduit şi să aşeze presbiteri în fiecare cetate\" (Tit 1:5). Tit a fost unul dintre constructorii edificiului Bisericii din primul veac. Ironic, numele pe care l-a purtat el a fost şi numele celui care în anul 70 d.Cr. a distrus complet Ierusalimul şi edificiul Templului: împăratul roman Titus.\r\n\r\nAutorul: Epistola este încă o lecţie de \"teologie pastorală\" pe care Pavel o dă tuturor păstorilor de-a lungul veacurilor prin intermediul destinatarului său imediat: Tit.\r\n\r\nData: Pavel scrie această epistolă în aceiaşi perioadă în care scrie şi cea dintîi epistolă către Timotei (63 d.Cr.), adică în intervalul de timp cuprins intre cele două detenţii în închisoarea Romei.\r\n\r\nContextul scrierii: Insula Creta este o fîsie de pămînt lungă de 230 de kilometri ieşită deasupra nivelului apelor acolo unde Marea Mediterană se uneşte cu Marea Egee. La vremea lui Pavel, cultura populaţiei de pe acele meleaguri era plină de mitologii şi curente filosofice păgîne. Tradiţional, insula fusese desemnată ca locul de naştere al lui Zeus, patronul panteonului grecesc, şi ca rezidenţă a minotaurului, jumătate om, jumătate bou, căruia regele Minos îi aducea sclavi ca jertfă de mîncare.\r\n\r\nEvanghelia s-a răspîndit repede printre locuitorii insulei şi adunările au apărut pretutindeni în oraşele insulei. Tit, unul dintre colaboratorii de încredere ai lui Pavel, a primit de la apostol sarcina, deloc uşoară, de a colinda Bisericile din oraşele Cretei şi de a supraveghea instalarea „presbiterilor\". Dincolo de aceasta, Pavel îi cere lui Tit să-i înveţe pe toţi credincioşii că fiecare are o lucrare de făcut pentru Domnul. Bărbaţi şi femei, tineri şi bătrîni, toţi trebuie să trăiască printre oameni ca o mărturie vie a credinţei lor creştine. Răspîndirea Evangheliei trebuia să fie făcută pe baza acestei mărturii colective a Bisericii.\r\n\r\nConţinutul cărţii: Epistola către Tit este foarte asemănătoare în conţinut cu prima epistolă trimisă lui Timotei: amîndouă prezintă rînduială care trebuie instaurată în Biserici, amîndouă prezintă caracterul şi caracteristicile pe care trebuie să le aibă cei ce vor să fie promovaţi ca presbiteri sau diaconi ai Bisericii şi tot amîndouă accentuează frumuseţea relaţiilor care trebuie să existe între membrii adunării creştine. Există însă o deosebire între cele două epistole: 1 Timotei scoate în relief „învăţătura\" despre organizarea Bisericii, iar Tit accentuează importanţa „faptelor\" membrilor Bisericii. Una vorbeşte despre „teorie\", iar cealaltă despre „practică\" (Tit 2:14; 3:8, 14).\r\n\r\nCuvinte cheie şi teme caracteristice: Versetul care rezumă elocvent întreaga epistolă este: „Adevărat este cuvîntul acesta, şi vreau să spui apăsat aceste lucruri, pentru ca cei ce au crezut în Dumnezeu, să caute să fie cei dintîi în fapte bune. Iată ce este bine şi de folos pentru oameni!\" (Tit 3:8).\r\n\r\nTeme mai importante din conţinutul cărţii sînt: calităţile şi responsabilităţile pastorale (Tit 1:5-9), norme de etică în viaţa celor credincioşi (Tit 2:1-10), revenirea Domnului Isus (Tit 2:11-14) şi procesul de mîntuire (Tit 3:3-7).\r\n\r\nCUPRINSUL CĂRŢII\r\n\r\nIntroducere, 1:1-4\r\n\r\nI. Pentru presbîterii Bisericii 1\r\na. Ca slujitori ai Bisericii - supraveghetori, 1:5, 6\r\nb. Ca oameni - fără prihană, 1:6-9\r\nc. Ca lucrători - sănătoşi în credinţă, 1:10-16\r\n\r\nII. Pentru categorii de credincioşi 2\r\na. Bărbaţii şi femeile în vîrstă, 2:2-3\r\nb. Slujirea femeilor în vîrstă, 2:4\r\nc. Pentru femeile tinere, 2:5\r\nd. Pentru tineri, 2:6\r\ne. Importanţa exemplului personal, 2:7-8\r\nf. Pentru cei ce sînt robi, 2:9-14\r\ng. Importanţa autorităţii apostolice, 2:15\r\n\r\nIII. Pentru toţi membrii Bisericii 3\r\na. Îndemn la fapte bune, 3:1 -2\r\nb. Faptele bune ca un rezultat al mîntuirii, 3:3-7\r\nc. Faptele bune ca o mărturie pentru alţii, 3:8\r\nd. Separarea de cei ce produc dezbinare, 3:9-11\r\n\r\nÎncheiere, 3:12-15" },
    // FILIMON
    { "FILIMON", "Explicatia Cartii: Filimon\r\n\r\nTitlul: În originalul grec, cartea poartă numele: „Pros Philemona\" - „Către Filimon\". Destinatarul epistolei este un om care reprezintă „paradoxurile\" în care trăiau creştinii primului secol. Filimon a fost un creştin stăpîn de sclavi. Evanghelia Domnului Isus este o forţă revoluţionară care trebuie să transforme oamenii şi relaţiile lor sociale. Cazul prezentat de epistola lui Pavel către Filimon este o ilustraţie a puterii de transformare exercitată de Evanghelie asupra celor ce o primesc şi i se supun.\r\n\r\nAutorul: Această epistolă este mai scurtă şi mai personală decît oricare altă scriere a lui Pavel. Ea nu tratează probleme de credinţă, ci caută să rezolve o criză apărută în relaţia dintre doi membrii ai Bisericii lui Cristos. Fără nici o îndoială, în afara epistolelor pe care le avem în Noul Testament, apostolul a mai scris multe alte scrisori ca aceasta. Filimon ne descopere latura pastorală a caracterului lui Pavel, talentul lui neasemuit de a se apropia de oameni şi abilitatea lui de a-i apropia pe oameni de Dumnezeu. Conţinutul epistolei nu prezintă nici un adevăr doctrinar major, dar textul este plin de tact, de delicateţe şi de bun simţ creştin. Maturitatea unui lucrător creştin se vede din felul în care el ştie să se descurce în situaţii neobişnuite şi complicate. Epistola către Filimon este o demonstraţie a maturităţii creştine atinse de Pavel. În ea observăm o impresionantă împletire de autoritate apostolică şi gingăşie pastorală.\r\n\r\nData: Epistola către Filimon este una din cele patru epistole scrise de Pavel din închisoarea Romei (vezi introducerea la epistola către Coloseni). Ea a fost trimisă în acelaşi timp cu epistolele scrise celor din Laodicea (Efeseni), Colose şi Filipi, adică în preajma anului 60 d.Cr. Versetul 22 ne spune că apostolul era închis, dar spera să fie eliberat în curînd: „pregăteşte-mi un loc de găzduire, căci trag nădejde să vă fiu dăruit, datorită rugăciunilor voastre\".\r\n\r\nContextul scrierii: Cadrul social din primul secol a pus multe şi chinuitoare probleme înaintea Evangheliei. Faptul că lumea era împărţită în sclavi şi oameni liberi şi că ţările se aflau sub stăpînirea nemiloasă a Romei a pus la grea încercare etica Bisericii creştine. Epistola către Filimon este un astfel de exemplu. Va putea dragostea propovăduită de creştinism să schimbe relaţiile dintre oameni? Va avea ea suficientă tărie pentru a împăca de pildă un stăpîn de sclavi cu sclavul fugar care se întoarce acasă? Răspunsul îl vom căpăta odată cu parcurgerea textului epistolei.\r\n\r\nConţinutul cărţii: Epistola către Filimon este o poveste de dragoste. Ea este corespondentul cărţii Rut din Vechiul Testament, cu deosebirea că în Rut avem de a face cu dragostea firească dintre oameni, iar în Filimon ne întîlnim cu dragoste „în Domnul\" care se manifestă între membrii colectivităţii creştine.\r\n\r\nEroii acţiunii sînt în număr de trei: Pavel, Filimon şi Onisim. La ceasul cînd scrie această scrisoare, Pavel se afla în închisoarea Romei. Dincolo de a fi preocupat de situaţia sa, el lucrează pentru Domnul şi continuă lucrarea care i-a fost atît de dragă: maturizarea unor caractere creştine în cei pe care i-a convertit prin Evanghelie. Acţiunea cărţii este simplă: Pe cînd era încă în Iudeea, Pavel a predicat în multe oraşe şi a trecut şi prin casa lui Filimon, care găzduia o adunare tînără de convertiţi la creştinism. Personalitatea lui Pavel a lăsat o puternică impresie asupra lui Filimon şi asupra celorlalţi din casa lui. La scurt timp după plecarea apostolului, din casa lui Filimon a fugit unul dintre sclavi: Onisim. Ca să-şi piardă urma, el s-a refugiat la Roma, unde putea trece neobservat în mulţimea pestriţă de locuitori. Prin coincidenţe pe care nu le cunoaştem, Onisim a ajuns să-l asculte pe Pavel în discuţiile pe care le purta cu evreii din Roma (Fapte 28:17-31) şi a fost convertit de Pavel la creştinism. După ce apostolul l-a ţinut pe lîngă el o vreme, Onisim a fost pus în faţa unui examen greu, care să-i încerce calitatea transformărilor pe care le-a lucrat Duhul lui Dumnezeu în inima lui. Ca o dovadă de pocăinţă, un creştin trebuie să caute să îndrepte răul pe care l-a săvîrşit înainte de convertire, aşa că apostolul Pavel i-a cerut lui Onisim să se întoarcă în Iudeea şi să accepte din nou sclavia în casa lui Filimon, răscumpărînd astfel „paguba\" produsă de fuga lui. Pavel nu-l trimite cu mîna goală, ci îi dă să ducă lui Filimon o scrisoare din partea sa. Aceasta este „epistola către Filimon\". Cine o citeşte cu atenţie, îşi dă repede seama că apostolul Pavel urmărea să „vîneze doi iepuri dintr-un singur foc\". Cazul lui Onisim este folosit şi pentru educarea sclavului fugit, dar şi pentru educarea stăpînului de sclavi în spiritul dragostei şi iertării creştine. Versetele 13 şi 14 ne arată că Pavel l-a trimis pe Onisim „doar ca să aibă de unde veni\": „Aş fi dorit să-l ţin la mine, ca să-mi slujească în locul tău cît sînt în lanţuri pentru Evanghelie. Dar n-am vrut să fac nimic fără învoirea ta, pentru ca binele, pe care mi-l faci, să nu fie silit, ci de bună voie\". Toată cheltuiala şi timpul pierdut cu transportul lui Onisim este suportată de apostol ca o investiţie în caracterul celor doi oameni. Citită din acest unghi, epistola îşi dezvăluie farmecul şi frumuseţea, asemeni unei flori care îşi deschide petalele.\r\n\r\nA reuşit Pavel ceea ce-şi propusese? L-a iertat Filimon pe Onisim şi a acceptat el să i-l trimită înapoi lui Pavel? Păstrarea acestei scrisori şi multiplicarea ei în copii care au ajuns pînă în zilele noastre este dovada clară că răspunsul la ambele întrebări a fost unul afirmativ. Cazul lui Onisim a fost cunoscut în toată comunitatea creştină şi a contribuit fără îndoială la rezolvarea multor altor conflicte apărute între fraţi din stări sociale şi materiale diferite. Urmarea a fost că Evanghelia a fărîmiţat bucăţică cu bucăţică imperiul Roman şi postulatele lui nedrepte de forţă pe care era aşezat, contribuind astfel la prăbuşirea lui mondială.\r\n\r\nCuvinte cheie şi teme caracteristice: Una din temele caracteristice acestei epistole este „triunghiul iubirii\". Pavel îi scrie lui Filimon cam aşa: Eu te iubesc pe tine şi ştiu că tu mă iubeşti pe mine, dar eu îl iubesc şi pe Onisim, aşa că dragostea ta pentru mine trebuie să se manifeste şi în atitudinea ta faţă de el („ca pe un frate prea iubit, mai ales de mine, şi cu atît mai mult de tine, fie în chip firesc, fie în Domnul! Dacă mă socoteşti dar ca prieten al tău, primeşte-l ca pe mine însumi\" v. 16-17). La scară umană, acest triunghi al iubirii este replica triunghiului dintre Dumnezeu şi oameni (1 Ioan 4:20-21)." },
    //EVREI
    { "EVREI", "Explicatia Cartii: Evrei\r\n\r\nEpistola către evrei răsare înaintea noastră ca un maiestuos vîrf de munte care domină toate culmile din depărtare. Ea este unul dintre cele două tratate de teologie sistematică din Noul Testament. Primul, epistola către Romani, a marcat intrarea în secţiunea dedicată epistolelor Bisericii dintre neamuri. Epistola către evrei marchează acum trecerea la cea de a doua secţiune a epistolelor creştine: epistolele Bisericii creştine a evreilor. Ceea ce urmează de aici înainte (Iacov, 1, 2 Petru, 1, 2, 3 Ioan, Iuda şi Apocalipsa) formează colecţia de epistole adresate prioritar evreilor. Ele sînt împreună „stîlpul evreiesc\" din edificiul arcului de triumf al credinţei creştine menţionat de noi în descrierea aşezării cărţilor Noului Testament.\r\n\r\nNiciuna dintre cele nouă epistole dedicate evreilor nu sînt adresate unei „biserici\", ci evreilor ca persoane particulare, ca grupuri distincte sau ca naţiune. Primul verset din epistola către evrei arată această schimbare de ton: „După ce a vorbit în vechime părinţilor noştri prin prooroci, în multe rînduri şi în multe chipuri (tipuri simbolice - n.n.), Dumnezeu...”\r\n\r\nAceasta nu înseamnă însă că cei dintre neamuri nu pot beneficia de pe urma citirii acestor epistole. Adevărurile cuprinse în ele sînt universal valabile, utile şi accesibile tuturor acelora care şi-au pus nădejdea în Domnul. Epistola către evrei, de pildă, ne arată clar supremaţia şi finalitatea revelaţiei mîntuitoare a lui Dumnezeu în Cristos. Nu există nici o singură altă cale de mîntuire. Nu există nimic care să poată fi pus alături de persoana şi lucrarea Mîntuitorului. La ce este şi la ceea ce a făcut El nu mai poate fi adăugat nimic şi din toate acestea nimic nu trebuie scos sau neglijat.\r\n\r\nTitlul: În original cartea poartă numele: „Pros Ebraious\" - „Către evrei\". Epistola este evreiască în temă, în conţinut şi în alcătuire. Terminologia folosită în text este aceea folosită în Sinagoga evreiască. Chiar şi numirea epistolei este făcută în vocabularul tipic sinagogii: „Vă rog, fraţilor, să primiţi bine acest cuvînt de sfătuire, căci v-am scris pe scurt\" (Evrei 13:22). Predica autorului, fiind scrisă, nu vorbită, poate fi asemuită unei părţi din „Midraş\"-ul evreilor creştini preocupat cu tălmăcirea creştină aplicată unor pasaje din Vechiul Testament şi mai ales din cartea Psalmilor. Numai în cuprinsul capitolului întîi găsim citate din Psalmul 2, 45, 102, 104 şi 110. În capitolul 2 sînt citate texte din Psalmul 8:4-6. În capitolul 3 este citat Psalmul 95:7-11. În capitolele 5 şi 6 găsim Psalmul 110:4, iar în capitolul 10 ne întîlnim cu pasaje din Psalmul 40:6-8. Unii au numit această epistolă: „Cea de-a cincea Evanghelie\". Primele patru descriu misiunea terestră a Domnului Isus, iar aceasta descrie misiunea Lui în cer, la dreapta Tatălui.\r\n\r\nAutorul: Deşi este atribuită lui Pavel, epistola care evrei nu-şi prezintă în nici un fel autorul şi nu ne dă nici un indiciu pentru identificarea lui. Împotriva părerii conform căruia Pavel ar fi autorul acestei cărţi este realitatea că Pavel a primit de la Domnul o misiune „pentru neamuri\". Discuţiile pentru stabilirea numelui celui care a scris această epistolă au continuat de-a lungul veacurilor pînă astăzi: Clement din Alexandria (150-215 d.Cr.) îl propune ca autor pe Pavel. Origen (185-253 d.Cr.) a fost de părere că gîndurile sînt ale lui Pavel, dar redactarea este a altui autor. Tertulian l-a sugerat pe Barnaba, Luther l-a propus pe Apolo, iar alţi comentatori au vorbit despre Filip, Evanghelistul sau despre Aquila şi Priscila. Consiliul întrunit la Cartagina în anul 397 d.Cr. îi atribuie lui Pavel scrierea a 14 epistole, printre care şi a aceleia scrisă evreilor. Probabil că şi noi va trebui să ne oprim la faimoasa remarcă a lui Origen: „Numai Dumnezeu însuşi ştie cine este autorul uman al epistolei\". Cît priveşte identitatea autorului divin, acesta este Dumnezeu însuşi.\r\n\r\nData: Clement al Romei citează din această epistolă într-una dintre lucrările sale, ceea ce face ca o dată după 96 d.Cr. să nu poată fi luată în consideraţie. Faptul că este citat sistemul aducerii jertfelor fără nici o aluzie la încetarea lui, ne conduce la concluzia că epistola a fost scrisă chiar şi înainte de căderea Ierusalimului şi distrugerea Templului (70 d.Cr.). Totuşi, destinatarii epistolei par a fi fost creştini deja de multă vreme (Evrei 5:12; 10:32-34), poate chiar creştini din a doua generaţie (Evrei 2 :13-14). Aceste observaţii fixează data scrierii epistolei undeva între anii 64-68 d.Cr.\r\n\r\nContextul scrierii: La data cînd a fost scrisă această epistolă, evreii, ca neam, îl refuzaseră de două ori pe Isus Cristos ca Mesia: prima dată cu ocazia răstignirii şi a doua oară după Rusalii. Totuşi, mulţimi mari de evrei au crezut în Domnul şi au format colectivităţi creştine compacte şi distincte de comunităţile religioase evreieşti grupate în jurul Templului şi al sinagogilor. Această „rămăşiţa\" care a primit noua revelaţie a lui Dumnezeu s-a aflat atacată deopotrivă din două părţi: din partea autorităţilor civile romane care s-au năpustit furibund asupra mişcării acestui „nou împărat: Isus\" (Fapte 17:7) şi din partea autorităţilor religioase tradiţionale evreieşti hotărîte să stăvilească orice dezvoltare ulterioară a acestei alternative spirituale numită: „Calea cea nouă\" (Fapte 13:25-26; 19:9; 24:14, 22; Rom. 7:6).\r\n\r\nConţinutul cărţii: Din cauza persecuţiei pornite împotriva creştinilor şi din cauza presiunii exercitate de concetăţenii lor religioşi, pentru mulţi evrei, părăsirea creştinismului şi întoarcerea la sistemul ritualistic iudaic părea o alternativă mai sigură şi mai comodă. Iată motivul pentru care autorul acestei cărţi îşi îndeamnă cititorii „să păstreze pînă la sfîrşit încrederea nezguduită şi nădejdea\" în Cristos (Evrei 3:6) şi să „meargă spre cele desăvîrşite\" (Evrei 6:1). Epistola către evrei are cel puţin trei scopuri precise:\r\n\r\na. Ea vrea să confirme valabilitatea creştinismului evreiesc prin evidenţierea faptului că venirea lui Isus Cristos a împlinit toate năzuinţele Iudaismului şi că în El au fost realizate toate profeţiile şi perceptele Legii din Vechiul Testament.\r\n\r\nb. Ea vrea să-i avertizeze pe evreii care au îmbrăţişat creştinismul asupra a două pericole: (1) pericolul întoarcerii la Iudaism şi (2) pericolul cochetării superficiale cu învăţătura creştină fără luarea unei hotărîri ferme şi definitive.\r\n\r\nc. Ea vrea să atragă atenţia creştinilor de pretutindeni asupra superiorităţii şi suveranităţii lui Cristos. Lucrarea Lui este superioară faţă de toate ritualurile şi instituţiile ceremoniale iudaice, iar persoana Lui este aşezată de Dumnezeu deasupra oricărei alte personalităţi sau oficialităţi religioase.\r\n\r\nCuvinte cheie şi teme caracteristice: Firul roşu care traversează toată cartea este ideea „superiorităţii lui Cristos\" (Evrei 1:4; 6:9; 7:7, 19, 22; 8:6; 9:23; 10:34; 11:16, 35, 40; 12:24). Epistola cuprinde o expunere a comparaţiei şi contrastului dintre lucrurile „bune\" ale Iudaismului şi lucrurile „mai bune\" aduse de Cristos. Domnul Isus este „mai bun\" decît îngerii, decît Moise, decît Iosua, decît Aaron; iar Legămîntul cel Nou este „mai bun\" decît Legămîntul mozaic (Evrei 8:7-13). Textul epistolei către evrei ni-L prezintă mai clar ca oriunde pe Dumnezeul-Om, Isus Cristos aşezat ca Mare Preot la dreapta măririi lui Dumnezeu şi mijlocind pentru mîntuirea oamenilor (Evrei 4:14-5:10; 6:20-8:13).\r\n\r\nMesajul întregii cărţi poate fi rezumat în conţinutul a două pasaje:\r\n\r\n„Astfei, fiindcă avem un Mare Preot însemnat, care a străbătut cerurile - pe Isus, Fiul lui Dumnezeu - să rămînem tari în mărturisirea noastră. Căci n-avem un Mare Preot care să n-aibă milă de slăbiciunile noastre; ci unul care în toate lucrurile a fost ispitit ca şi noi, dar fără păcat. Să ne apropiem dar cu deplină încredere de scaunul harului, ca să căpătăm îndurare şi să găsim har, pentru ca să fim ajutaţi la vreme de nevoie\" (Evrei 4:14-16)\r\n\r\n„Şi noi, dar, fiindcă sîntem înconjuraţi cu un nor aşa de mare de martori, să dăm la o parte orice piedică, şi păcatul care ne înfăşoară aşa de lesne, şi să alergăm ai stăruinţă în alergarea care ne stă înainte\" (Evrei 12:l5.\r\n\r\nCUPRINSUL CĂRŢII\r\n\r\nI. MESAGERUL „MAI BUN\": FIUL\r\na. Superioritatea fata de profeţi, 1:1-3\r\nb. Superioritatea fată de îngeri, 1:4-14\r\nParanteză: Pericolul neglijării, 2:1-4\r\nc. Întruparea, 2:5-18\r\n\r\nII. APOSTOLUL „MAI BUN\"\r\na. Superioritatea fată de Moise, 3:1-6\r\nParanteză: Pericolul necredinţei, 3:7-19\r\nb. Superioritatea persoanei Sale, 4:1-10\r\nParanteză: Pericolul neascultării, 4:11-13\r\n\r\nIII. PREOTUL „MAI BUN\"\r\na. Comparaţia cu Aaron, 4:14-5:4\r\nb. Rînduială lui Melhisedec, 5:5-7:25\r\nRînduit, 5:5-6\r\nAutorul mîntuirii, 5:7-10\r\nParanteză: Pericolul imaturităţii, 5:11-6:12\r\nÎnainte mergătorul, 6:13-20\r\nUn preot viu, 7:1-17\r\nÎntărit prin jurămînt, 7:18-25\r\nc. Relaţia cu jertfele, 7:26-28\r\n\r\nIV. LEGĂMÎNTUL „MAI BUN\"\r\na. Stabilirea legămîntului, 8:1-13\r\nb. Conţinutul vechiului legămînt, 9:1-10\r\nc. Cristos şi Noul Legămînt, 9:11 -28\r\n\r\nV. JERTFA „MAI BUNĂ\"\r\na. Neputinţa Legii, 10:1-4\r\nb. Jertfa lui Cristos, 10:5-18\r\nParanteză: Pericolul respingerii, 10:19-31\r\n\r\nVI. CALEA „MAI BUNĂ\": CREDINŢA\r\na. Necesitatea credinţei, 10:32-39\r\nb. Exemple de credinţă, 11:1-40\r\nc. Exersarea credinţei, 12:1-17\r\nd. Obiectivul credinţei. 12:18-24\r\nParanteză: Pericolul refuzului, 12:25-29\r\n\r\nVII. PRACTICAREA CREDINŢEI\r\na. În relaţiile sociale, 13:1 -6\r\nb. În relaţiile spirituale, 13:7-17\r\nSalutări personale, 13:18-25\r\n\r\nSCHIŢĂ TEMATICĂ\r\n\r\nI. ISUS - Un eliberator „mai bun\"\r\na. Isus Omul-Dumnezeu - mai bun ca îngerii\r\nb. Isus Noul Apostol - mai bun ca Moise\r\nc. Isus Noua Căpetenie - mai bun ca Iosua\r\nd. Isus Noul Preot - mai bun decît Aaron\r\n\r\nII. GOLGOTA - Un Legămînt „mai bun\"\r\na. are promisiuni mai bune\r\nb. descinde dintr-un Cort mai bun\r\nc. este pecetluit cu o jertfă mai bună\r\nd. aduce rezultate mult mai bune\r\n\r\nIII. CREDINŢA - Calea mai bună\r\na. este răspunsul cerut de Dumnezeu\r\nb. a fost calea aleşilor lui Dumnezeu\r\nc. trebuie să privească acum spre Domnul\r\nd. este arătată prin trăirea în sfinţenie\r\n\r\nCuvînt de încheiere" },
    //IACOV
    { "IACOV", "Explicatia Cartii: Iacov\r\n\r\nTitlul: În original, cartea poartă numele: „Iakobou Epistole\" - „Epistola lui Iacov\".\r\n\r\nAutorul: În cuprinsul Noului Testament întîlnim trei persoane care au purtat acest nume: (1) Iacov, fiul lui Zebedei şi fratele lui Ioan, care a fost din numărul celor 12 apostoli şi care a devenit primul apostol martir în anul 44 d.Cr. (despre el citim în Mat. 4:21; 10:2; 17:1, Luca 5:10; Fapte 12:1-2), (2) Iacov, fiul lui Alfeu, şi el unul dintre cei 12, dar despre el nu ştim nici un fel de detalii (Mat. 10:3; Marcu 3:18; Luca 6:15; Fapte 1:13) şi (3) Iacov, unul dintre cei patru fraţi mai mici ai Domnului Isus (Matei 13:55; Marcu 6:3). Acest Iacov a fost la început un obstacol în calea oamenilor către Isus (Matei 13:55), apoi a căutat să-L oprească pe Domnul din activitatea Lui (Matei 12:46-50). Acestea s-au întîmplat pentru că Iacov n-a crezut în dumnezeirea lui Isus (Ioan 7:5). După înviere, Domnul Isus i s-a arătat în mod special (1 Cor. 15:7) convingîndu-l pe deplin, alipindu-l de grupul celorlalţi apostoli (Fapte 1:14) şi rînduindu-l să fie promovat în fruntea Bisericii din Ierusalim, alături de Ioan şi Petru (Fapte 12:17; 15:13-29; 21:17-18; Gal. 1:19; 2:9, 12; Iuda 1). Toate evidenţele îl indică pe acest al treilea Iacov drept autor al epistolei.\r\n\r\nData: Iosif Flavius, un istoric evreu, scrie că Iacov, fratele Domnului a fost martirizat în anul 62 d.Cr., aşa că trebuie să plasăm data scrierii epistolei ceva mai devreme. Lipsa oricăror aluzii la controversele doctrinare discutate la Consiliul de la Ierusalim, ne îndreptăţeşte să plasăm data scrierii probabil undeva între anii 48-50 d.Cr. Dacă aşa stau lucrurile, atunci avem de a face cu cea mai timpurie scriere creştină dintre toate cele care s-au păstrat pînă în zilele noastre.\r\n\r\nContextul scrierii: În calitatea sa de presbiter al bisericii din Ierusalim, Iacov scrie această epistolă către: „cele doisprezece seminţii care sînt împrăştiate\" (Iacov 1:1). Expresia folosită identifică grupurile de evrei care trăiau în afara hotarelor Palestinei. Convertiţii lui Petru din ziua de Rusalii fuseseră şi ei „iudei, oameni cucernici din toate neamurile care sînt sub cer\" (Fapte 2:5). Fără nici o îndoială că aceşti noi creştini s-au înapoiat în ţinuturile lor şi au dus în comunităţile lor vestea despre lucrarea lui Isus Mesia. Iacov scrie pentru membrii Bisericii care se află în tranziţia dinspre Iudaismul apostolilor spre universalitatea Evangheliei vestite de Pavel. A spune însă că accentul pe care-l pune Iacov pe importanţa faptelor este o încercare de corectare a învăţăturii lui Pavel înseamnă a face o mare greşală. La ora aceea, Pavel nu-şi scrisese nici epistola către Romani şi nici epistola către Galateni. Asupra evreilor convertiţi la creştinism se dezlănţuise persecuţia şi prigoana. De aceea, Iacov îşi începe epistola îndemnîndu-i să reziste în „feluritele încercări\" (Iacov l:2) şi o termină sfătuindu-i să fie „îndelung răbdători\": „Fiţi şi voi îndelung răbdători, întăriţi-vă inimile, căci venirea Domnului este aproape\" (Iacov 5:7, 8)\r\n\r\nConţinutul cărţii: Textul are cinci aspecte caracteristice: (1) nu există nici o referire la creştinii dintre neamuri sau la relaţia dintre creştinii evrei şi creştinii proveniţi dintre alte popoare, aşa cum găsim în epistolele scrise la o dată mai tîrzie, (2) în afară de faptul că este pomenit numele Domnului Isus, textul nu cuprinde practic nici o dezbatere sau dizertaţie teologică, ceea ce ne trimite iarăşi la o dată timpurie, cînd creştinismul era considerat numai un fel de Iudaism mesianic, (3) aluziile la învăţăturile lăsate de Domnul Isus sînt atît de sărace încît ne îndeamnă să credem că această epistolă a fost scrisă chiar înainte de publicarea Evangheliilor, (4) Iacov foloseşte termenul de „sinagogă\" alături de acela de Biserică (Iacov 2:2; 5:14) ceea ce arată că evreii creştini erau organizaţi după tiparul simplu al rînduielilor aşezămintelor de învăţătură iudaice (Iacov 3:1; 5:14), şi (5) Iacov nu aminteşte, în nici un fel, dezbaterile sau hotărîrile luate la Consiliul de la Ierusalim (anul 49 d.Cr.)\r\n\r\nConţinutul epistolei poate fi grupat sub tema: „Credinţa adevărată este o credinţă militantă, care se manifestă prin fapte\". Ideile majore ale epistolei sînt următoarele: credinţa ne ajută să biruim toate încercările vieţii (capitolul 1), credinţa ne ajută să arătăm aceiaşi bunăvoinţă faţă de toţi oamenii (capitolul 2), credinţa ne transformă şi inima şi felul nostru de vorbire (capitolul 3), credinţa ne învaţă să trăim cu evlavie în toate aspectele vieţii (capitolul 4) şi această credinţă ne face să aşteptăm venirea Domnului Isus ca pe o rezolvare a tuturor suferinţelor şi necazurilor (capitolul 5).\r\n\r\nCuvinte cheie şi teme caracteristice: Epistola lui Iacov a fost înţeleasă greşit de mulţi şi considerată de o mai mică valoare spirituală. Verificarea calităţii credinţei prin mărturia faptelor exterioare rămîne însă, oricît nu ne-ar place, singura omologare acceptată de Dumnezeu: „După roadele lor îi veţi cunoaşte\" (Mat. 7:16-20).\r\n\r\nFaptul că Pavel spune că Avraam a fost socotit neprihănit prin credinţă, iar Iacov afirmă că acelaşi Avraam a fost socotit neprihănit prin fapte (Iacov 2:21) nu este nici un fel de contradicţie. Pavel şi Iacov se referă la două evenimente succesive din viaţa lui Avraam. Socotit neprihănit prin credinţă atunci cînd „L-a crezut pe Dumnezeu pe cuvînt\" şi a ieşit din ţara lui, Avraam şi-a dovedit apoi calitatea credinţei sale prin faptul că L-a ascultat pe Dumnezeu şi a fost gata să-l aducă pe Isaac ca jertfă (Iacov 2:21). Înaintea lui Dumnezeu, credinţa este cea care justifică omul, iar faptele sînt cele care omologhează credinţa. Chiar Iacov subliniază această dublă verificare atunci cînd, după versetul omologării credinţei lui Avraam prin fapte, citează versetul la care se va opri mai tîrziu Pavel: „Astfel s-a împlinit Scriptura care zice: „Avraam a crezut pe Dumnezeu, şi i s-a socotit ca neprihănire\" (Iacov 2:23).\r\n\r\nUnii au încercat să spună că ar exista în textul lui Iacov două feluri de credinţe: una cu fapte şi una fără fapte. O astfel de interpretare îl nedreptăţeşte pe autorul acestei epistole. Iacov nu face deosebire între două feluri de credinţe, ci între credinţa vie şi credinţa moartă, adică inexistentă.\r\n\r\nEpistola lui Iacov rămîne şi astăzi o „oglindă\" în care ne putem analiza calitatea credinţei noastre (Iacov l:22-24). Ea trebuie citită cel puţin din timp în timp, pentru a ne feri de ipocrizie şi de formalismul religios, gol şi lipsit de viaţă. O schiţă â cărţii este greu de făcut şi nu este neapărat necesară." },
    // 1 PETRU
    { "1 PETRU", "Explicatia Cartii: 1 Petru\r\n\r\n„V-am scris ca să vă adeveresc că adevăratul har al lui Dumnezeu este harul acesta, de care v-aţi alipit\"\r\n\r\nTitlul: În original, epistola începe cu cuvintele: „Petros apostolos Iesou Christou\" - „Petru, apostol al lui Isus Cristos\". De aici se trage şi numele ei: „Petrou A\" - „Petru A\".\r\n\r\nAutorul: Epistola este incontestabil un produs al lui Petru, apostolul Domnului, fratele lui Andrei şi fiul lui Iona (Mat. 16:17). Locul său de naştere a fost Betsaida, sat de pescari pe malul mării Galileii (Ioan l:41-42). Petru a fost unul dintre cei trei ucenici care au format anturajul intim al Domnului Isus (Marcu 5:37; 9:2; 14:33). În cîteva ocazii, Petru s-a bucurat de o atenţie specială din partea Mîntuitorului (Luca 5:10; Matei 16:17; Luca 22:31-32; Ioan 13:6-10). După înviere şi Rusalii, Petru a devenit purtător de cuvînt pentru grupul apostolic. El a călătorit intens, vizitînd Bisericile şi exercitînd asupra lor autoritatea apostolică. Toate scrierile şi cuvîntările sale sînt pline de autoritate şi de înţelepciune. Mai ştim despre Petru că a fost căsătorit şi că a fost însoţit în călătoriile sale de soţia sa (1 Cor. 9:5). Unde nu a ajuns să meargă personal, apostolul a trimis scrisori pastorale. Se pare că această primă epistolă a fost trimisă prin Silvanus (1 Petru 5:12). Acest colaborator apostolic a mai făcut astfel de servicii şi pentru Pavel (2 Cor. 1:19; 1 Tes. 1:1; 2 Tes. 1:1).\r\n\r\nData: Epistola a fost probabil scrisă în anul 64 d.Cr., cu puţin timp înainte de izbucnirea prigoanelor lui Nero împotriva creştinilor.\r\n\r\nContextul scrierii: Viaţa lui Petru a suferit o schimbare dramatică după învierea Domnului Isus, iar persoana sa a ajuns să ocupe un loc proeminent în Biserica primară. Misiunea lui Petru a fost îndreptată în mod special înspre poporul evreu. Iată cum clarifică apostolul Pavel această situaţie:\r\n\r\n„...mie îmi fusese încredinţată Evanghelia pentru cei netăiaţi împrejur, după cum lui Petru îi fusese încredinţată Evanghelia pentru cei tăiaţi împrejur, - căci Cel ce făcuse din Petru apostolul celor tăiaţi împrejur, făcuse şi din mine apostolul neamurilor - şi cînd au cunoscut harul care-mi fusese dat, Iacov, Chifa şi Ioan, care sînt priviţi ca stîlpi, mi-au dat mie şi lui Barnaba, mîna dreaptă de însoţire, ca să mergem să propovăduim: noi la neamuri, iar ei la cei tăiaţi împrejur\" (Gal. 2:7-10).\r\n\r\nDupă ce citim cuvintele lui Pavel, înţelegem foarte clar de ce epistolele lui Iacov, Petru şi Ioan sînt grupate separat într-o secţiune dedicată scrierilor îndreptate prioritar către creştinii iudei. De fapt, Petru ne spune el însuşi în debutul epistolei sale că le scrie: „aleşilor care trăiesc ca străini, împrăştiaţi prin Pont, Galatia, Capadocia, Asia şi Bitinia\" (1 Petru 1:1).\r\n\r\nNu este foarte clar dacă Petru s-a aflat la data scrierii în Babilonul de pe rîul Eufrat (1 Petru 5:12) sau dacă această numire este o metaforă sub care este ascunsă identitatea Romei, în orice caz, acest Petru devenise între timp unul dintre conducătorii grupului apostolic. El a fost purtătorul lor de cuvînt în ziua de Rusalii (Fapte 2) şi asupra activităţii lui este concentrată atenţia primelor 12 capitole din cartea Faptelor Apostolilor. După ce Pavel a preluat misiunea cu Evanghelia între neamuri, Petru a rămas să misioneze printre evrei, dar aceasta nu înseamnă că el nu s-a ocupat şi de cei proveniţi din alte popoare. De fapt, majoritatea noilor Biserici creştine formate erau deja grupări mixte, în care deosebirile dintre iudei şi neamuri trecuseră pe planul al doilea.\r\n\r\nConţinutul cărţii: Petru îşi trimite scrisoarea către creştinii care trăiau ca „străini\" într-o lume din ce în ce mai ostilă Bisericii. În mijlocul nemurilor înfuriate şi al evreilor fanatici, creştinii începuseră să sufere din cauza ataşamentului lor faţă de Cristos. Petru le scrie pentru a-i întări în credinţă. Apostolul îi încurajează să se poarte într-un chip demn de persoana şi lucrarea Mîntuitorului. Fiind născuţi prin credinţă la o nădejde nouă, ei sînt sfătuiţi să urmeze pilda lui Cristos. Petru le spune că asemănarea lor cu Cristos trebuie să se materializeze în două domenii: credinţa lor trebuie să-i conducă la o viaţă de supunere şi la o viaţă de acceptare a suferinţei. Ca cetăţeni, ei trebuie să fie supuşi autorităţilor, ca robi, ei trebuie să le fie supuşi stăpînilor lor, ca soţi şi soţii şi ca membrii în adunare, ei trebuie să fie supuşi unii altora. În 1 Petru, credinţa îi face pe cei credincioşi să „se supună\" (1 Petru 2:13-19; 3:l-7), să „sufere\" (1 Petru 2:19-21; 3:14, 17; 4:1, 12:16) şi să „aştepte venirea Domnului\" (1 Petru 1:3, 13, 21; 3:15; 4:13; 5:14).\r\n\r\nCuvinte cheie şi teme caracteristice: Duhul Sfînt a rînduit epistolele în grupuri semnificative. Am văzut cum după Credinţa din evrei au urmat Faptele din Iacov ; acum vom vedea cum Petru ca autor de scrieri inspirate se distinge ca un apostol al nădejdii, tot aşa cum Pavel a fost un apostol al credinţei, iar Ioan a fost un apostol al iubirii. Cuvîntul „nădejde\" apare în 1 Petru 1:3, 13, 21; 3:15.\r\n\r\nSCHIŢA CĂRŢII\r\n\r\nIntroducere, 1:1-2\r\n\r\nI. NĂDEJDEA CEA VIE\r\na. nădejdea vie pusă în practică, 1:3-12\r\nb. cuvîntul viu pus în practică, 1:12-2:3\r\nc. piatra cea vie şi poziţia noastră, 2:4-10\r\n\r\nII. VIAŢA DE MĂRTURIE\r\na. ca cetăţeni, 2:12-17\r\nb. ca robi, 2:18-25\r\nc. ca soţi şi soţii, 3:1-7\r\nd. ca străini între oameni, 3:8-4:6\r\ne. în relaţiile dintre membrii Bisericii, 4:7-11\r\n\r\nIII. „ÎNCERCAREA DE FOC\"\r\na. „Bucurie\" şi „încredere\" în încercare, 4:12-19\r\nb. Presbiteri credincioşi în slujbă, 5:1-4\r\nc. Toţi să trăiască în smerenie şi aşteptare, 5:5-11\r\n\r\nÎncheiere, 5:12-14" },
    // 2 PETRU
    { "2 PETRU", "Explicatia Cartii: 2 Petru\r\n\r\nTitlul: În originalul grec, cartea poartă numele: „Petrou B\" -„Petru B\".\r\n\r\nAutorul: Primul verset al cărţii îl prezintă pe autor drept „Simon Petru, rob şi apostol al lui Isus Cristos\" (2 Petru 1:1). Numele de „Simon\" este o aducere aminte a vieţii lui Petru dinainte de întîlnirea cu Domnul Isus, „Petru\" este numele pe care l-a primit acest apostol prin Cristos. „Petru\" înseamnă „stîncă\" şi desemnează simbolic stabilitatea şi statornicia. Viaţa şi scrisorile apostolului sînt marcate din plin de aceste caracteristici. 2 Petru 3:1 ne spune clar că epistola este o continuare a mesajului din prima scrisoare: „Prea iubiţilor, aceasta este a doua epistolă pe care v-o scriu. În amîndouă caut să vă trezesc mintea sănătoasă prin înştiinţări\".\r\n\r\nData: Cea de a doua epistolă a fost scrisă la puţin timp după cea dintîi, probabil din acelaşi loc. Pentru mai multe detalii vă rugăm să citiţi introducerea făcută celei dintîi epistole a lui Petru.\r\n\r\nContextul scrierii: Cea de a doua epistolă a lui Petru este o chemare la seriozitate şi la curăţie. 1 Petru s-a ocupat cu problemele care au asaltat Biserica din afară. 2 Petru tratează problemele care pot măcina viaţa Bisericii din lăuntru. Apostolul le scrie credincioşilor ca să-i avertizeze de pericolul „învăţătorilor mincinoşi\" strecuraţi în rîndul credincioşilor. El începe prin a le atrage tuturor atenţia asupra vieţii lor personale de umblare cu Domnul. Vieţuirea creştină presupune perseverentă şi sîrguinţă în credinţă şi fapte, în cunoştinţă şi înfrînare, în răbdare şi evlavie, în dragoste de fraţi şi în iubire de oameni. Prin contrast cu acestea, învăţătorii mincinoşi sînt dedaţi plăcerilor, obraznici, pofticioşi şi lacomi. Ei batjocoresc venirea Domnului şi Judecata viitoare, lansîndu-se în desfrîuri şi petreceri. Petru vrea să le aducă aminte tuturor că deşi este îndelung răbdător, Domnul îşi va împlini planurile cu pămîntul şi va răsplăti fiecăruia după faptele lui. Cine ştie aceasta, face bine dacă trăieşte frumos şi în curăţie, pregătindu-se în fiecare zi a călătoriei lui înspre lucrurile viitoare.\r\n\r\nConţinutul cărţii: Aşa cum a fost arătat deja, cea de a doua epistolă a lui Petru este un avertisment împotriva lucrării proorocilor mincinoşi: „În norod s-au ridicat şi prooroci mincinoşi, cum şi între voi vor fi învăţători mincinoşi, care vor strecura pe furiş erezii nimicitoare, se vor lepăda de Stăpînul, care i-a răscumpărat, şi vor face să cadă asupra lor o pierzare năpraznică. Mulţi îi vor urma în destrăbălările lor. Şi, din pricina lor, calea adevărului va fi vorbită de rău\" (2 Petru l-2).\r\n\r\nAceastă a doua espistolă seamănă foarte mult cu cea de a doua epistolă scrisă de Pavel lui Timotei. Şi Petru, ca şi Pavel, se aşează la scris cu sentimentul că viaţa lui se apropie foarte repede de sfîrşit: „Dar socotesc că este drept, cît voi mai fi în cortul acesta, să vă ţin treji aducîndu-vă aminte; căci ştiu că dezbrăcarea de cortul meu va veni deodată, după cum mi-a arătat Domnul nostru Isus Cristos\" (2 Petru 1:14; Ioan 21:18-19). Amîndouă epistole sînt luminoase, chiar dacă întrevăd viitoarea lepădare de credinţă şi decadenţa care va caracteriza „zilele din urmă\". Secretul optimismului autorilor lor este în faptul că amîndoi priveau dincolo de orizontul timpului, spre revenirea Domnului Isus şi spre încoronarea Lui în slavă.\r\n\r\nTeme importante din cuprinsul epistolei lui Petru sînt: perseverenţa sfinţilor ca un răspuns dat „alegerii divine\" (2 Petru l:4-14), aducerea aminte despre „schimbarea la faţă\" petrecută cu Domnul pe munte (2 Petru 1:15-18), învăţătura despre inspirarea şi tălmăcirea Scripturilor (2 Petru 1:19-21), învăţătura despre venirea Domnului (2 Petru 3:4-13, precum şi îndemnurile la vigilenţă şi credincioşie (2 Petru 3:14-17).\r\n\r\nCuvinte cheie şi teme caracteristice: Tema centrală a epistolei este „cunoaşterea\". Nimic nu este mai important într-o vreme de rătăciri spirituale decît o cunoaştere adecvată a Scripturilor. Verbul „a cunoaşte\" şi derivatele lui apar de 13 ori în textul scrisorii. O altă caracteristică a acestei epistole este caracterul ei escatologic. Aflat el însuşi în preajma morţii, Petru sfătuieşte Biserica să se încreadă în promisiunile Domnului şi să creadă că El îşi va duce la bun sfîrşit programul său cu lumea şi că în curînd vom păşi sub un cer noii pe un „pămînt în care va locui neprihănirea\" (2 Petru l:21; 3:10-13).\r\n\r\nSCHIŢA CĂRŢII\r\n\r\nIntroducere, 1:1\r\n\r\nI. Îndemnare apostolică\r\na. Promisiuni scumpe, 1:2-4\r\nb. Progres spre ţintă, 1:5-7\r\nc. Priorităţi sfinte, 1:8-11\r\n\r\nII. Mărturie apostolică\r\na. Mărturie despre adevăr, 1:12-15\r\nb. Trăire în adevăr, 1:16-18\r\nc. Studiu despre adevăr, 1:19-21\r\n\r\nIII. Avertizare apostolică\r\na. Învăţătura proorocilor mincinoşi, 2:1-3\r\nb. Exemple de prooroci mincinoşi, 2:4-9\r\nc. Caracterul proorocilor mincinoşi, 2:10-19\r\nd. Soarta proorocilor mincinoşi, 2:20-22\r\n\r\nIV. Nădejdea apostolică\r\na. Cei ce se îndoiesc de promisiuni, 3:1-7\r\nb. Certitudinea împlinirilor profetice, 3:8-10\r\nc. Vieţuirea în aşteptarea marilor împliniri, 3:11-18" },
    // 1 IOAN
    { "1 IOAN", "Explicatia Cartii: 1 Ioan\r\n\r\nTitlul: Cu toate că numele autorului nu apare nicăieri în text, numele acestei epistole este în originalul grec: ”Ioanou A\" - „Ioan A\", ceea ce o face cea dintîi epistolă dintre cele trei ale lui Ioan.\r\n\r\nAutorul: Între Evanghelia lui Ioan şi aceste trei epistole atribuite lui există o identitate de stil care nu poate fi contestată de nimeni. Apostolul Ioan a fost fiul lui Zebedei şi fratele lui Iacov, cel dintîi martir al Bisericii creştine. Ioan şi Iacov au fost împreună cu Petru în cercul „celor trei\" pe care Domnul Isus i-a luat pretutindeni cu Sine. La Cruce, Domnul Isus i-a încredinţat lui Ioan îngrijirea mamei Sale (Ioan 19:26). După înviere, înălţare şi Rusalii, Ioan a devenit unul din stîlpii spirituali ai Bisericii din Ierusalim.\r\n\r\nData: Ioan a scris această epistolă către Bisericile din Asia Mică, amintite şi în Apocalipsa. Probabil că data scrierii a fost undeva între anii 85-95 d.Cr. Fiind mai tînăr, Ioan a supravieţuit tuturor celorlalţi apostoli şi a ajuns să fie privit ca apărător al credinţei creştine într-o vreme cînd „ereziile\" atacau crezul Bisericii creştine.\r\n\r\nContextul scrierii: Fără nici o îndoială, studiul cărţilor Bibliei, şi al epistolelor „ioanine\" în special, trebuieşte făcut şi cu ochiul şi inima păstorului, căci pe lîngă datele statistice, istorice sau stilistice pe care le înregistrează mintea exegetului, textul biblic mai prezintă şi un mesaj personal, „către credincioşi\".\r\n\r\nFără a nega deosebitul conţinut teologic al acestor scrieri, dorim să subliniem existenţa unui mesaj preponderent personal, adresat unei adunări particulare, (sau unui grup din adunare), care se afla într-o anumită situaţie:\r\n\r\n„Ei au ieşit din mijlocul nostru, dar nu erau dintre ai noştri. Căci dacă ar fi fost dintre ai noştri, ar fi rămas cu noi; ci, ei au ieşit ca să se arate că nu toţi sînt dintre ai noştri\" (1 Ioan 2:19).\r\n\r\nPretutindeni forma de adresare este: „eu\", „voi\", „noi\" obişnuită în conversaţia celor care se cunosc, iar destinatarii epistolelor sînt numiţi „copilaşi preaiubiţi\". Autorul îi iubeşte pe cei cărora le scrie. Este profund preocupat de protejarea lor faţă de influenţele lumii şi faţă de ereziile falşilor învăţători. Doreşte creşterea lor spirituală în dragoste, credinţă şi sfinţenie. Pentru realizarea acestor deziderate, el face mereu apel la ceea ce ei sînt şi la ceea ce ei cunosc. Îi îmbărbătează, îi mustră, polemizează cu ei, şi îi învaţă. Toate aceste amănunte sînt particulare activităţii unui păstor şi se regăsesc astăzi în viaţa acelora pe care acelaşi Mare Păstor i-a chemat, învrednicindu-i să le încredinţeze o parte a turmei Sale.\r\n\r\nFără nici o îndoială, epistolele au şi un caracter polemic. Ele nu sînt în nici un caz tratate teologice scrise în liniştea academică a vreunei biblioteci, ci mesaje fierbinţi, izvorîte din necesitatea stringentă a rezolvării problemelor care se iviseră. Dintre toate aceste probleme care apăruseră în Biserică, cea dintîi epistolă a lui Ioan se ocupă cu propaganda insidioasă a unor falşi învăţători: „Copilaşilor, nimeni să nu vă înşele\" (1 Ioan3:7), „V-am scris aceste lucruri în vederea celor ce cautăsă vă rătăcească\" (1 Ioan 2:26).\r\n\r\nDupă unii comentatori erezia din Biserica la care se referă apostolul Ioan poate fi încadrată în erezia „Docetismului\". Numele acestui sistem derivă de la verbul grecesc „dokein\" - „a părea\", „a fi aparent\". În dogmatica lor Isus „părea\" a fi om, era uman numai „în aparenţă\", căci în El era o teofanie asemănătoare cu celelalte „arătări\" din Vechiul Testament, care se iveau oridecîteori Dumnezeu sau îngerul Domnului se descopereau oamenilor sub formă umană.\r\n\r\nAstăzi noi cunoaştem această erezie din scrierile „patristice\" în care „părinţii bisericeşti\" au folosit chiar această primă epistolă a lui Ioan pentru a o combate. Printre ei i-am putea cita pe Ignatius, Polycarp şi Tertulian.\r\n\r\nUn studiu atent al frazeologiei lui Ioan ne va arăta însă că erezia din Biserica de pe timpul lui Ioan nu era legată aşa de mult de realitatea trupului lui Cristos, ci mai mult de relaţia dintre aspectul uman al lui „Isus\" şi aspectul divin al „Fiului\" şi al „Cristosului\". Accentul negaţiei nu se pune prea mult pe umanitatea reală a lui Isus, ci pe identitatea „Cristosului pre-existent\" cu „Omul Isus\".\r\n\r\nAcest lucru i-a condus pe majoritatea comentatorilor la concluzia că ereticii pomeniţi în epistole pot şi trebuie să fie încadraţi în gruparea „gnosticilor\".\r\n\r\nSă încercăm o scurtă prezentare a acestui sistem. „Gnosticismul\" este un termen generic care cuprinde de fapt mai multe sisteme dogmatice păgîne, dar şi evreieşti sau chiar pseudo-creştine: un fel de sincretism filozofico-religios prin care se creeau bazele unui sistem universal valabil, accesibil tuturor oamenilor, indiferent de climatul spiritual în care s-au născut.\r\n\r\nLa origine, gnosticismul a fost o învăţătură păgînă care a reuşit să combine în sine elemente ale intelectualismului occidental cu fondul mistic propriu orientului. Plummer rezumă acest sistem la două coordonate fundamentale: „impuritatea materiei\" şi „supremaţia conştiinţei\". Într-adevăr preocuparea de bază a acestor „gnostici\" era tocmai eliberarea spiritului de trup, pe care-l priveau numai ca pe o „închisoare materială a spiritului\".\r\n\r\nConceptul de „materie coruptă definitiv şi păcătoasă în structură\" a fost comun ambelor sisteme de religie, occidental şi oriental. Părerea aceasta a dat naştere unei teorii despre existenţa unei succesiuni valorice de „eoni\" sau emanaţii din Fiinţa Supremă. Aceste sfere concentrice succesive de „eoni\" se aflau la distanţe din ce în ce mai mari de „pleroma\" sau mediul existenţei de Sine a Fiinţei Supreme, valoarea lor divină scăzînd direct proporţional cu îndepărtarea de sursa de sfinţenie. La periferia tuturor, ultima din ierarhia valorică a fiinţării, s-ar fi aflat lumea materială.\r\n\r\nA. Trebuie să remarcăm faptul că polemica pe care o duce Ioan cu ei are ca obiect „întruparea\", posibilitatea ca Dumnezeul absolut în sfinţenie să se dezbrace de slavă şi să ia chip de om. Într-adevăr religia creştină susţine că Fiul lui Dumnezeu s-a întrupat El însuşi şi că trupul fiecărui credincios devine un „templu sfînt\" al Duhului lui Dumnezeu.\r\n\r\nDeparte de a putea accepta o mîntuire prin „trupul jertfit de Cristos\", gnosticii erau adepţii unei mîntuiri prin „iluminarea spiritului\". Această „iluminare\" se putea produce printr-o cunoaştere „esoterică\" însuşită în cadrul unor ceremonii speciale. Iniţiaţii deveneau „psuchikoi\", oameni care şi-au trezit puterile latente ale sufletului, ridicîndu-se deasupra muritorilor de rînd.\r\n\r\nCea mai timpurie tradiţie asociază Epistolele lui Ioan cu viaţa Bisericii din Asia Mică. Subscriem şi noi acestei păreri cu atît mai mult cu cît se ştie că „gnosticismul\" se infiltrase mai ales în mişcarea creştină din acele regiuni, iar prezenţa unui „mare\" iniţiat ca Cerintius din Efes nu putea rămîne fără un răspuns public din partea apostolului Ioan.\r\n\r\nDespre Cerintius, Irineu ne spune că el „susţinea că Isus nu ar fi fost născut dintr-o fecioară, ci ar fi fost fiul natural al Mariei şi al lui Iosif, lucru care bineînţeles că nu l-a împiedicat de fel să devină cel mai drept şi mai înţelept dintre toţi oamenii timpului său. Mai tîrziu, mult mai tîrziu, după botezul său, Cristosul a coborît asupra acestui om normal, sub forma unui porumbel, simbol al trimisului de la Supremul Stăpîn. Din acea clipă, omul Isus a început să-L propovăduiască pe Tatăl „cel necunoscut\" şi să înfăptuiască minuni. La sfîrşitul vieţii lui Isus, „divinul\" Cristos l-a părăsit şi l-a lăsat singur să fie prins, să sufere, şi să moară. Toate aceste evenimente penibile au fost trăite numai de pămîntescul Isus, în timp ce Cristosul s-a detaşat impasibil, senin şi de neatins ca orice fiinţă spirituală\".\r\n\r\nÎn esenţă, erezia lui Cerintius consta în această separare categorică a omului Isus de Cristosul divin (sau Duhul) emanat iniţial şi apoi reîntors neatins în „pleroma\".\r\n\r\nUn cititor atent va remarca imediat că textul epistolelor lui Ioan conţine cîteva expuneri de argumente îndreptate tocmai împotriva ereziilor lui Cerintius.\r\n\r\nFără a-l mai numi în mod expres în text, Ioan pare a se referi la el în 1 Ioan 2:22: „Cine este mincinosul, dacă nu cel ce tăgăduieşte că Isus este Cristosul? Acela este Anticristul, care tăgăduieşte pe Tatăl şi pe Fiul\". Tot aşa în 1 Ioan 4:3 şi în 2 Ioan 7: „Duhul lui Dumnezeu să-L cunoaşteţi după aceasta: Orice duh, care mărturiseşte că Isus Cristos a venit în trup, este de la Dumnezeu; şi orice duh, care nu mărturiseşte pe Isus, nu este de la Dumnezeu, ci este duhul lui Anticrist, de a cărui venire aţi auzit. El chiar este în lume acum\", „Căci în lume s-au răspîndit mulţi amăgitori, care nu mărturisesc să Isus Cristos vine în trup. Iată amăgitorul, iată Anticristul!\"\r\n\r\nIoan scrie în capitolul 5:6 că: „El, Isus Cristos, este Cel ce a venit cu apă şi cu sînge; nu numai cu apă (referire la epifania de la botez), ci cu apă şi cu sînge\" (sîngele se referă la suferinţele şi moartea Sa).\r\n\r\nCu alte cuvinte, spre deosebire de Cerintius care susţinea că „Cristosul divin\" s-a pogorît asupra lui Isus după botez şi că L-a părăsit înainte de moarte, Ioan accentuează faptul că Isus Cristos, persoană unică şi unitară a trecut şi prin botez şi prin moarte.\r\n\r\nB. Despre latura morală a sistemului „gnostic\" găsim informaţii în scrierile lui Irineus şi Eusebius.\r\n\r\nConform mărturiilor acestor doi autori, erezia morală a lui Cerintius a fost: „...împreunarea între bărbaţi... mult timp cunoscută sub numele de Nicolaitism\". Despre Nicolaiţi citim şi în Apocalipsa 2:6, 14, 15 unde sînt amintiţi ca unii „cu fapte şi cu o învăţătură\" pe care Dumnezeu „le urăşte\".\r\n\r\nAcestei depravări morale, Ioan îi dă răspuns în 1 Ioan 3:3, 9 şi mai ales în 1 Ioan 3:6: „Oricine are nădejdea aceasta în el, se curăţeste, după cum El este curat\", „Oricine este născut din Dumnezeu nu păcătuieşte... pentru că este născut din Dumnezeu\" şi „Oricine rămîne în El, nu păcătuieşte; oricine păcătuieşte nu L-a văzut, nici nu L-a cunoscut\".\r\n\r\nPentru Ioan, naşterea din Dumnezeu şi trăirea într-o practică împotriva ordinii stabilite de Dumnezeu prin creaţie sînt incompatibile şi ireconciliabile.\r\n\r\nC. A treia caracteristică a „gnosticilor\", incluzîndu-l bineînţeles şi pe Cerintius, pare a fi fost totala lor lipsă de dragoste. Proclamîndu-se un fel de „aristocraţie\" a iluminaţilor, singurii care au ajuns să cunoască „adîncimile\", gnosticii erau stăpîniţi de un profund dispreţ faţă de ceilalţi oameni şi în primul rînd faţă de creştini.\r\n\r\nIoan răspunde acestei învăţături astfel: „Cine zice că este în lumină, şi urăşte pe fratele său, este încă în întunerec pînă acum\" (1 Ioan 2:9).\r\n\r\nCuvîntul „dragoste\" este folosit de atîtea ori în epistolele lui Ioan încît creştinii din toate timpurile au ajuns să-l supranumească „apostolul iubirii\". Se pare că Ioan şi-a meritat cu prisosinţă acest frumos nume. Într-unul din comentariile sale la Galateni 6:10, Ieronim ne spune ceva caracteristic despre viaţa „binecuvîntatului Ioan Evanghelistul\": „Ajuns la o vîrstă înaintată în Efes, apostolul era de obicei purtat pe braţe şi adus în mijlocul adunării, căreia îi repeta mereu unul şi acelaşi mesaj al dragostei: „Copilaşilor, iubiţi-vă unii pe alţii\". Cîteodată ei îl întrebau: „Învăţătorule, de ce ne spui mereu asta?\" Răspunsul lui era invariabil acesta: „Pentru că este porunca Domnului şi pentru că dacă o împliniţi numai pe aceasta este deajuns...”\r\n\r\nConcluzia tuturor celor spuse pînă aici este că împotriva ereziilor despre persoana lui Isus Cristos, împotriva indiferentismului moral şi împotriva aroganţei lipsite de dragoste a gnosticilor cerintieni, apostolul Ioan îşi clădeşte răspunsul pe trei stîlpi de forţă ai adevăratului creştinism: credinţa în Isus ca şi Cristos întrupat, ascultarea de poruncile Domnului şi trăirea în dragostea frăţească.\r\n\r\nConţinutul cărţii: Ioan încearcă să-i păzească pe cei din Biserică de primejdia acestor „falşi învăţători\", care tulburau viaţa credincioşilor. El atrage atenţia asupra a trei caracteristici ale lor: originea lor diabolică, influenţa lor drăcească şi învăţătura lor falsă. Iată de ce el îi numeşte:\r\n\r\n1. „Falşi profeţi\" (1 Ioan 4:1). Un profeţeşte un om care vorbeşte sub directa influenţă a unei puteri supranaturale. Profetul adevărat este „gura\" prin care vorbeşte Dumnezeu. Profetul mincinos este dimpotrivă „gura\" prin care se face auzit „duhul rătăcirii\". Iată de ce examinarea învăţăturii unui profet este sinonimă cu o „cercetare a duhurilor\" (1 Ioan 4:1-6).\r\n\r\n2. „Amăgitori\" (2 Ioan 7). Prin aceasta apostolul îi indentifică cu acei care duc poporul în rătăcire, promiţîndu-le lucruri care nu există de fapt.\r\n\r\n3. „Anticrişti\" (1 Ioan 2:18 cf.v. 22; 4:3; 2 Ioan 7). Mesajul învăţăturii lor neagă personalitatea divin-umană a Domnului Isus Cristos.\r\n\r\nTrebuia să remarcăm faptul că de fiecare dată apostolul Ioan ne spune că aceste persoane nedorite în colectivitatea Bisericii erau „mulţi\"; „mulţi falşi profeţi\", „mulţi amăgitori\" şi „mulţi anticrişti\". Acest lucru ne este confirmat şi de faptul că ei au reuşit pentru o perioadă de timp să treacă drept membrii ai Bisericii. În momentul scrierii epistolei dintîi această fracţiune „ne-creştină\" se separase de Biserică: „au ieşit din mijlocul nostru\" şi „s-au dus în lume\" (1 Ioan 2:19; 2 Ioan 7). Totuşi aceasta ruptură a reuşit să-i tulbure pe mulţi dintre membrii adunării, aşa că Ioan s-a văzut nevoit să le scrie. El îi laudă pe cei care prin rămînerea lor în Cristos „i-au biruit\" pe duşmanii adevărului (1 Ioan 4:4).\r\n\r\nCuvinte cheie şi teme caracteristice: Tema principală a acestei epistole este: „Certitudinea creştină\". Verbele cele mai folosite în textul ei sînt: „ginoskein\" - a cunoaşte, a observa, a pricepe (de 15 ori), şi „eidenai\" - a fi sigur pe ceea ce şti. Cuvîntul caracteristic acestei epistole este: „parresia\" -„îndrăzneală, încredere în atitudine\".\r\n\r\nCertitudinea creştină este o realitate cu două aspecte, (a) unul obiectiv (pentru că religia creştină este o sumă de adevăruri), şi (b) unul subiectiv (pentru că fiecare creştin a fost „născut din nou\" şi are în sine „arvuna vieţii veşnice\". Epistola lui Ioan este tocmai o expunere şi o argumentare a acestor două aspecte ale certitudinii creştine.\r\n\r\nCunoaşterea creştină este o cunoaştere absolută şi duce la o certitudine absolută. Bineînţeles că este vorba despre aspectul calitativ al cunoaşterii, nu de cel cantitativ. Creştinul este cel care cunoaşte „adevărul\" lucrurilor. El cunoaşte adevărul despre lume şi despre starea ei (1 Ioan 5:19, 2:18, 3:15), adevărul despre el însuşi, despre datoria şi destinul lui (1 Ioan 2:10, 11, 29; 3:2; 5:18), şi mai presus de toate, creştinul cunoaşte adevărul despre Dumnezeu şi despre Isus Cristos, (1 Ioan 5:20; 2:13, 14; 4:6, 7).\r\n\r\nCa să ajungem la certitudinea creştină despre persoana şi lucrarea lui Cristos avem în primul rînd:\r\n\r\n1. Evenimentul istoric. Domnul Isus Cristos a fost „trimis\" (1 Ioan 4:9, 10, 14). El „a venit\" (1 Ioan 5:20) şi „s-a manifestat\" său „a fost arătat\" („ephanerothe\" în 1 Ioan l:2; 3:5, 8; 4:9). Această venire a Lui a fost „în trup\" (1 Ioan 4:2; 2 Ioan 7), cu „apă şi cu sînge\" (1 Ioan 5:6).\r\n\r\nToate aceste realităţi îl obiectivizează şi-L îmbracă în modul cel mai absolut cu experienţa naşterii, cu botezul şi cu trăirea morţii. Evenimentul în sine nu a putut trece neobservat. Cel care a venit în „trup\" a trăit printre oameni şi prin relatările lor ajungem la a doua sursă a cunoaşterii:\r\n\r\n2. Mărturia apostolilor - „Ce era de la început, ce am auzit, ce am văzut cu ochii noştri, ce am privit şi am pipăit cu mîinile noastre, cu privire la Cuvîntul vieţii - pentru că viaţa a fost arătată, şi noi am văzut-o, şi mărturisim despre ea, şi vă vestim viaţa veşnică, viaţă care era la Tatăl, şi care ne-a fost arătată; - deci, ce am văzut şi auzit, aceea vă vestim şi vouă, ca şi voi să aveţi părtăşie cu noi. Şi părtăşia noastră este cu Tatăl şi cu Fiul Său, Isus Cristos\", „Şi noi am văzut şi mărturisim că Tatăl a trimis pe Fiul casă fie Mîntuitorul lumii\" (1 Ioan 1:1-3; 4:14).\r\n\r\nAl treilea lucru care ne dă certitudine este:\r\n\r\n3. „Ungerea Duhului Sfînt\". Ea funcţionează ca un al şaselea simţ prin care putem pătrunde toate lucrurile (1 Corint. 2:10) ca să ajungem la „gnosis\", la cunoaştere: „Cît despre voi, ungerea pe care aţi primit-o de la El, rămîne în voi, şi n-aveţi trebuinţă să vă înveţe cineva; ci, după cum ungerea Lui vă învaţă despre toate lucrurile şi este adevărată, şi nu este o minciună, rămîneţi în El, după cum v-a învăţat ea\" (1 Ioan 2:20, 27 cf. 3:2; 4:13).\r\n\r\nAceastă cale de cunoaştere este în lăuntrul creştinului şi se întregeşte cu dovezile exterioare ale „apei şi sîngelui\" (1 Ioan 5:6, 8, 9).\r\n\r\nÎn demonstraţia pe care o face pentru a-i convinge pe credincioşi despre certitudinea vieţii veşnice, Ioan este cel puţin tot atît de preocupat să aducă argumente care să dovedească faptul că cei care nu cred în Fiul lui Dumnezeu nu au viaţa veşnică, oricît de „iluminaţi\" ar fi ei.\r\n\r\nDistincţia aceasta între adevăraţii creştini pe care vrea să-i întărească şi adversarii eretici cu care se înfruntă este vizibilă în tot conţinutul epistolei. Pretutindeni întîlnim „voi\" şi „ei\":\r\n\r\nAceste două grupări distinctive există şi astăzi. Unii, încrezuţi şi plini... de ceea ce de fapt nici nu posedă, iar ceilalţi, frecventatori din obicei ai bisericilor, care nu au siguranţa mîntuirii şi cărora li se pare chiar o obrăznicie să susţii că aşa ceva poate exista! Toţi aceştia trebuie să afle că există o siguranţă creştină, o veritabilă certitudine care nu este nici arogantă, nici înşelătoare, ci dimpotrivă luminoasă şi cu prisosinţă revelată de însuşi Dumnezeu.\r\n\r\nPentru a confirma şi cerceta calitatea credinţei, apostolul Ioan le propune creştinilor trei teste caracteristice:\r\n\r\n1. Testul teologic (teoretic). Prin acest test se verifică mesajul credinţei noastre. Adevărata credinţă susţine că Isus este „Fiul lui Dumnezeu\" (1 Ioan 3:23, 5:5, 10, 12, 20) şi că „Cristosul a venit în trup\" (1 Ioan 4:2, 6; 2 Ioan 7).\r\n\r\nNici un sistem de dogme sau doctrine care neagă preexistenta eternă a lui Isus sau încarnarea lui istorică nu poate fi acceptat ca fiind creştin: „Oricine tăgăduieşte pe Fiul, n-are pe Tatăl\" (2:23).\r\n\r\n2. Testul moral. Testul acesta trebuie să verifice dacă noi trăim în neprihănire şi în păzirea poruncilor lui Dumnezeu. În epistola lui Ioan, păcatul este arătat a fi total incompatibil cu natura lui Dumnezeu, care este din acest punct de vedere definit prin „lumină\" (1:5). Păcatul este un accident nedorit în lumea lui Dumnezeu, de aceea Ioan ne spune că Fiul lui Dumnezeu „s-a arătat ca să ia (să înlăture) păcatele; şi în El nu este păcat\" (3:5). Concluzia limpede care reiese din aceste două afirmaţii este aceea că oricine este născut „din Dumnezeu\"„nu păcătuieşte\" pentru că a pus capăt unei vieţi de păcat (1 Ioan 3:9).\r\n\r\nOrice experienţă „mistică\", presupusă a fi creştină, însoţită de imoralitate trebuie imediat abandonată: „Dacă zicem că avem părtăşie cu El, şi umblăm în întuneric, minţim şi nu trăim adevărul\" (1:6).\r\n\r\n3. Testul social. Cel de al treilea test verifică atitudinea noastră faţă de ceilalţi. „Dumnezeu este dragoste\" postulează Ioan, aşa că toţi cei născuţi din El trebuie să moştenească această aplecare plină de afecţiune faţă de ceilalţi oameni.\r\n\r\n„Prea iubiţilor, să ne iubim unii pe ceilalţi; căci dragostea este de la Dumnezeu. Şi oricine iubeşte, este născut din Dumnezeu, şi cunoaşte pe Dumnezeu. Cine nu iubeşte, n-a cunoscut pe Dumnezeu, pentru că Dumnezeu este dragoste\" (1 Ioan 4:7, 8).\r\n\r\nA nu fi în stare să treci aceste trei teste ale certitudinii despre viaţa veşnică înseamnă a nu fi de fapt copil al lui Dumnezeu:\r\n\r\n„Dacă zicem că avem părtăşie cu El, şi umblăm în întuneric, minţim şi nu trăim adevărul\" (1:6).\r\n\r\n„Cine zice: „Îl cunosc\", şi nu păzeşte poruncile Lui, este un mincinos, şi adevărul nu este în el\" (2:4).\r\n\r\n„Cine zice că este în lumină, şi urăşte pe fratele lui, este încă în întuneric pînă acum\" (2:9).\r\n\r\nO certitudine solidă despre Cristos şi despre viaţa veşnică este singura forţă care poate anima mărturisirea creştină a Bisericii.\r\n\r\nSCHIŢA CĂRŢII\r\n\r\nPrefaţa (1:1-4)\r\n\r\nI. Mesajul apostolic şi implicaţiile lui morale (1:5 - 2:2)\r\nÎmpotriva:\r\na. negării faptului că păcatul rupe părtăşia noastră cu Dumnezeu (1:6, 7)\r\nb. negării faptului că păcatul există în însăşi natura noastră (1:8, 9)\r\nc. negării faptului că păcatul se arată pe sine în purtarea noastra (1:10- 2:2)\r\n\r\nII. Prima aplicare a testelor (2:3-27)\r\na. ascultarea sau „testul moaral\" (2:3-6)\r\nb. dragostea sau „testul social\" (2:7-11)\r\nc. o digresiune despre Biserică (2:13-14)\r\nd. o digresiune despre lume (2:15-17)\r\ne. credinţa sau „testul doctrinal\" (2:18-27)\r\n\r\nIII. A doua aplicare a testelor (2:28 - 4:6)\r\na. o detailare a testului moral = neprihănirea (2:28 - 3:10)\r\nb. o detailare a testului social = dragostea (3:11-18)\r\nc. o digresiune despre siguranţă şi despre îndoială (3:19-24)\r\nd. o detailare a testului doctrinal = credinţa (4:1-6)\r\n\r\nIV. A treia aplicare a testelor (4:7 - 5:5)\r\na. o nouă dezvoltare a testului social = dragostea (4:7-12)\r\nb. o combinare a testelor doctrinal şi social (4:13-21)\r\nc. o combinare a celor trei teste (5:1-5)\r\n\r\nV. Cei trei martori şi siguranţa pe care ne-o dau ei (5:6-17)\r\na. cei trei martori (5:6-12)\r\nb. siguranţa noastră ca o consecinţă (5:13-17)\r\n\r\nVI. Trei afirmaţii şi o scurtă îndrumare (5:18-21)\r\na. „Ştim\" că neprihănirea este divină (5:18)\r\nb. „Ştim \" că lumea zace în păcat (5:19)\r\nc. „Ştim\" că Evanghelia este mîntuitoare (5:20)\r\nd. De aceea să ne păzim de tot ceea ce nu este de la Dumnezeu (5:21)" },
    // 2 IOAN
    { "2 IOAN", "Explicatia Cartii: 2 Ioan\r\n\r\nCele trei epistole ale lui Ioan sînt adresate în ordine: unei Biserici, unei familii şi unei persoane. Cea de a doua epistolă este singura scriere din Biblie adresată unei mame creştine.\r\n\r\nTitlul: În originalul grec, cartea poartă numele: „Ioannou B\" - „Ioan B\" spre a fi deosebită de cea dintîi epistolă a lui Ioan.\r\n\r\nAutorul: Fără nici o îndoială că „Presbiterul\" amintit la începutul epistolei este apostolul Ioan. Limbajul folosit este acela al autorului celei de a patra Evanghelii. Numai el goate folosi un fond atît de restrîns de cuvinte şi totuşi să exprime o gamă atît de variată de adevăruri. Într-adevăr, Ioan foloseşte cele mai simple cuvinte cu putinţă; vocabularul lui este acela al unui copil la vîrsta de 5-7 ani, dar conţinutul de informaţii, bogăţia de idei şi imagini sînt fascinant de bogate. Apostolul scrie despre: adevăr, lumină, dragoste, umblare, rămînere, viaţă, apă, ură şi moarte. Însemnătatea acestor cuvinte trece însă cu mult peste folosul lor obişnuit, devenind ferestre spre nişte realităţi spirituale profunde.\r\n\r\nData: Epistola a fost scrisă probabil în preajma anului 90 d.Cr.\r\n\r\nContextul scrierii: Este clar că cea de a doua epistolă trebuie aşezată în aceleaşi circumstanţe spirituale şi istorice ca şi prima epistolă. Este interesant să remarcăm că Duhul Sfînt ne-a lăsat un instructaj complet de comportament în condiţii de atacuri asupra Bisericii. În 1 Ioan ni s-a spus cum trebuie să se comporte adunarea ca întreg, în 2 Ioan ni s-a spus cum trebuie să se comporte fiecare familie din adunare, iar în 3 Ioan ni s-a spus cum trebuie să se comporte fiecare credincios în parte.\r\n\r\nCei ce încearcă să spiritualizeze conţinutul acestei epistole spun că ar fi adresată unei Biserici pe care autorul o numeşte conspirativ: „aleasă Doamnă şi copiii ei\" (2 Ioan 1:1). O citire atentă a versetului 4 şi mai ales 10-13 ne va arăta însă că aceasta este o tălmăcire forţată. Ce fel de Biserică ar fi aceea în care numai „unii din copiii tăi umblă în adevăr\"?\r\n\r\nCasa acestei mame creştine căreia îi scrie apostolul era o casă creştină. Ioan nu pierde ocazia să-şi exprime bucuria pentru atmosfera în care erau crescuţi copiii acestei femei. „Umblarea\" lor era „în adevăr\", dar pericolul era şi el pe aproape. Proorocii mincinoşi dădeau tîrcoale celor ce mergeau pe calea dreaptă şi căutau să se furişeze în casele lor şi să le strecoare în suflet veninul învăţăturilor lor drăceşti. În condiţii normale, ospitalitatea este o înaltă virtute creştină, dar în contactele cu „amăgitorii\" (1 Ioan 7), ospitalitatea se poate dovedi primul pas spre dezastru. Aşa cum remarca cineva: „Minciuna este un misionar plin de rîvnă. Ea merge din casă în casă căutînd să convertească cît mai multe persoane\". Cu maturitatea care-l caracteriza, apostolul Ioan o sfătuieşte pe această mamă creştină să-şi păzească cu vigilenţă „cuibul\".\r\n\r\nConţinutul cărţii: În cea dintîi epistolă, Ioan ne-a informat că un grup de oameni din Biserică părăsiseră adunarea şi învăţătura creştină: „Ei au ieşit din mijlocul nostru, dar nu erau dintre ai noştri. Căci dacă ar fi fost dintre ai noştri, ar fi rămas cu noi; ci ei au ieşit ca să se arate că nu toţi sînt dintre ai noştri\" (1 Ioan 2:19). Ieşiţi din adunare, aceşti „amăgitori\" colindau din casă în casă, infiltrîndu-se prin Biserici şi oferind creştinilor „cunoştinţe mai înalte decît Evanghelia\", propovăduită în Biserică. Apostolul Ioan o avertizează pe această mamă creştină să nu se lase tîrîtă în erezie şi să-i refuze categoric pe cei ce vor încerca să o depărteze de Domnul: „Păziţi-vă bine să nu vă pierdeţi rodul muncii voastre, ci să primiţi o răsplată deplină. Oricine o ia înainte, şi nu rămîne în învăţătura lui Cristos, n-are pe Dumnezeu\" (2 Ioan 8-9). Dragostea de oameni este o virtute, compromisul cu ereticii este un păcat, iar părtăşia cu batjocoritorii Domnului este la fel de vinovată ca şi păcătuirea.\r\n\r\nCuvinte cheie şi teme caracteristice: Cuvîntul „adevăr\" apare de cinci ori în primele 4 versete. Întreaga epistolă este un avertisment împotriva asaltului minciunii şi împotriva imposturii celor care colindau din casă în casă ca să „buimăcească familii întregi, învăţînd pe oameni, pentru un cîştig mîrşav, lucruri pe care nu trebuie să le înveţe\" (Tit 1:11). Tema generală a epistolei este „statornicia în Evanghelia Domnului\" şi este enunţată în versetul 6: „Şi dragostea stă în vieţuirea după poruncile Lui. Aceasta este porunca în care trebuie să umblaţi, după cum aţi auzit de la început\".\r\n\r\nSCHIŢA CĂRŢII\r\n\r\nIntroducere, 1-3\r\n\r\nI. ASCULTAREA DE PORUNCA DOMNULUI\r\na. Umblarea în adevăr, 4-5\r\nb. Umblarea în dragoste, 6\r\n\r\nII. VEGHERE ÎN ASCULTARE DE DOMNUL\r\na. Avertisment împotriva falşilor învăţături 7-9\r\nb. Avertisment împotriva falsei ospitalităţi, 10-11\r\n\r\nÎncheiere, 12-13" },
    // 3 IOAN
    { "3 IOAN", "Explicatia Cartii: 3 Ioan\r\n\r\nTitlul: În originalul grec, cartea poartă numele: „Ioannou G\" - „Ioan G\", „gama\" fiind a treia literă din alfabetul grecesc.\r\n\r\nAutorul: Acelaşi „presbiter\" care a scris şi cea de a doua epistolă este şi autorul acesteia (2 Ioan 1; 3 Ioan 1). De data aceasta, apostolul Ioan se adresează nu Bisericii şi nici unei familii ci unui frate din Biserică.\r\n\r\nData: A doua şi a treia epistolă a lui Ioan au fost probabil scrisori de salut care au însoţit cea dintîi epistolă spre destinatarii ei. Data scrierii lui 3 Ioan deci trebuie să fie plasată tot în preajma anului 90 d.Cr.\r\n\r\nContextul scrierii: Cea de a treia epistolă s-a născut pe fondul aceloraşi frămîntări care tulburau viaţa Bisericii creştine din Asia spre sfîrşitul vieţii lui Ioan. Rămas ultimul ucenic al Domnului în viaţă, Ioan veghează asupra adunărilor creştine şi asupra celor care funcţionează ca lideri spirituali ai Bisericii.\r\n\r\nConţinutul cărţii: Această a treia epistolă a lui Ioan este scurtă în conţinut, dar plină de informaţii care trebuiesc toate studiate. Epistola poate fi considerată o scrisoare de însoţire şi recomandare pentru „cărăuşii\" trimişi să răspîndească epistola scrisă de Ioan pentru apărarea credinţei în acele timpuri de atacul furibund al ereticilor „gnostici\": „Vei face bine să îngrijeşti de călătoria lor, într-un chip vrednic de Dumnezeu; căci au plecat pentru dragostea Numelui Lui, fără să primească ceva de la neamuri. Este datoria noastră dar, să primim bine pe astfel de oameni, ca să lucrăm împreună cu adevărul\" (3 Ioan 7-8).\r\n\r\nLa fel de bine însă, epistola poate fi citită ca şi un studiu în comportamentul liderilor spirituali ai Bisericii. Ioan, Gaiu, Diotref şi Dimitrie sînt tot atîtea tipuri de slujitori ai Bisericii. Ioan este „presbiterul\" cu autoritate apostolică, Gaiu este lucrătorul tînăr plin de rîvnă care se avîntă dincolo de limitele puterilor sale (3 Ioan 2-3). Creşterea lui spirituală era supravegheată direct de Ioan. În Biserica în care se afla, Gaiu este prins Între exemplul rău dat de Diotref („Diotref, căruia îi place să aibă întîietatea între ei, nu vrea să ştie de noi. Ne cleveteşte cu vorbe rele, nu primeşte pe fraţi şi împiedică şi pe cei ce voiesc să-i primească, şi-i dă afară din Biserică\" - v. 9-10) şi exemplul bun dat de Dimitrie („Toţi, chiar şi Adevărul, mărturisesc bine despre Dimitrie; şi noi mărturisim despre el; şi ştii că mărturisirea noastră este adevărată\" - v.12). Sfatul pe care Ioan i-l dă lui Gaiu este să nu se lase biruit de rău şi să se ia după exemplul bun stabilit de Dimitrie: „Prea iubitule, nu urma răul, ci binele. Cine face binele este de la Dumnezeu; cine face răul n-a văzut pe Dumnezeu\" - v. 11).\r\n\r\nCuvinte cheie şi teme caracteristice: În textul acestei epitole ne întîlnim iarăşi cu „umblarea în adevăr” (3 Ioan 4), cu „umblarea în dragoste\" (3 Ioan l, 6, 7) şi cu grija lui Ioan pentru Biserica Domnului. Intimitatea dintre apostol şi credincioşi este ilustrată cum nu se poate mai bine de exprimarea dorinţei lui Ioan de a sta de vorbă cu ei „gura\" către gură\" (2 Ioan 12; 3 Ioan 13).\r\n\r\nSCHIŢA CĂRŢII\r\n\r\nIntroducere, 1\r\n\r\nI. GAIU - LUCRAREA ÎN ADEVĂR ŞI ÎN DRAGOSTE\r\nUn lucrător în creştere\r\na. Credincioşie faţă de adevăr, 3-4\r\nb. Slujire fată de fraţi, 5\r\nc. Umblare în dragoste, 6\r\nd. Colaborare cu fraţii, 7-8\r\n\r\nII. DIOTREF, LUCRAREA ÎN FIREA PĂMÎNTEASCĂ\r\nUn lider firesc\r\na. Îi place să aibă întîietatea, 9\r\nb. Nu vrea să ştie de alţii, 9\r\nc. Îi cleveteşte pe alţi lucrători, 10\r\nd. Nu primeşte pe fraţi, 10\r\ne. Crede că este„proprietarul\" Bisericii, 10\r\nf. Dă afară pe cine vrea el, 10\r\ng. Va trebui să fie disciplinat, 10\r\nh. Este un pericol pentru creşterea altora, 10\r\n\r\nIII. DIMITRIE, LUCRAREA VORBITĂ DE BINE\r\nUn exemplu demn de urmat\r\na. Sprijinit pe Cuvîntul lui Dumnezeu, 12,\r\nb. Vorbit de bine de fraţii din Biserică, 12\r\nc. Confirmat de apostoli, 12\r\n\r\nÎncheiere, 13-14" },
    // IUDA
    { "IUDA", "Explicatia Cartii: Iuda\r\n\r\nTitlul: Cartea se numeşte în original: „Iouda\" - „Iuda\", după numele celui ce a scris-o.\r\n\r\nAutorul: Numele celui care a scris-o este dat chiar în conţinutul cărţii (v.1). Acest Iuda a fost un alt frate al Domnului Isus (Mat. 13:55; Marcu 6:3). El nu se consideră în numărul celor 12 apostoli (Iuda 17), ci se prezintă ca „frate al lui Iacov\" (Iuda 1). În mod obişnuit, în vremea aceea o persoană se identifica pe sine după numele tatălui său. Motivul pentru care Iuda a făcut excepţie de la regula aceasta a fost probabil dublu: (1) din modestie şi respect, el n-a vrut să facă aluzie la relaţia lui de familie cu Isus Cristos, şi (2) el crede că poate fi identificat foarte bine în funcţie de relaţia cu Iacov, fratele său mai mare care ajunsese între timp unul dintre liderii proeminenţi în Biserica din Ierusalim.\r\n\r\nData: Epistola a fost scrisă probabil cîndva în perioada cuprinsă între anii 65-80 d.Cr. Erezia pe care o combate Iuda a apărut în Biserica primară destul de repede şi a fost combătută cu putere de Pavel, Ioan şi Petru. De fapt, epistola lui Iuda poate fi considerată ca o reluare a celei de a doua epistole a lui Petru. Între 2 Petru 2:1-22 şi Iuda 4-18 există o asemănare imposibil de trecut cu vederea. Probabil că Iuda a avut de confruntat aceleaşi probleme ca şi Petru şi s-a folosit de epistola şi autoritatea apostolului pentru a-şi întări şi mai mult punctul de vedere.\r\n\r\nContextul scrierii: În timp ce Iuda se pregătea să le scrie fraţilor despre mîntuire, el se vede silit să-şi schimbe subiectul pentru a combate activitatea şi învăţătura unui grup de oameni plini de vicii, care circulau prin Biserici şi căutau „să schimbe în desfrînare harul lui Dumnezeu\" (Iuda 3-4). Nu este greu să recunoaştem în acest grup de prooroci mincinoşi pe „amăgitorii\" pe care-i prevestise Petru (2 Petru 2:1) şi pe care-i combătuse cu atîta putere Ioan (vezi 1 Ioan). Aparent aceşti învăţători mincinoşi căutau să-i convingă pe cei credincioşi că harul iertării se întinde nu numai în trecut, dar şi în prezent. Ei considerau mîntuirea ca pe un fel de paşaport spre lumea trăirii în pofte şi păcate. Ei negau dumnezeirea lui Cristos şi reduceau creştinismul la o sumă de cunoştinţe teoretice fără legătură cu viaţa de toate zilele.\r\n\r\nConţinutul cărţii: Cine citeşte această epistolă îşi dă repede seama că Iuda scrie prin excelenţă evreilor. Mulţimea de citate şi exemple din Vechiul Testament se succed fără nici o lămurire suplimentară, indicînd faptul că Iuda presupunea că cititorii lui sînt de mult familiarizaţi cu istoria Israelului şi cu conţinutul Scripturilor.\r\n\r\nIuda le aminteşte creştinilor felul în care s-a purtat Dumnezeu în trecut cu necredinciosul Israel, cu îngerii neascultători, cu oraşele păcătoase Sodoma şi Gomora şi cu aceeia care, asemenea lui Cain, Balaam şi Core, s-au răzvrătit împotriva Domnului (Iuda 5, 6, 7, 8-10, 11). După un scurt pasaj în care descrie imoralitatea acestor prooroci mincinoşi, Iuda încheie scurta lui epistolă printr-un avertisment fierbinte, printr-o chemare urgentă la statornicie (Iuda 12-19, 20-23) şi printr-una din cele mai frumoase benedicţii din Biblie (Iuda 24-25).\r\n\r\nSCHIŢA CĂRŢII\r\n\r\nIntroducere, 1-2\r\n\r\nI. ATACUL DUŞMANILOR\r\na. Urgenţa apelului din epistolă, 3\r\nb. Doctrina învăţătorilor mincinoşi\r\n- schimbă harul în desfrînare, 4\r\n- tăgăduiesc dumnezeirea lui Cristos, 4\r\n\r\nc. Soarta învăţătorilor mincinoşi Exemplul\r\nÎngerilor căzuţi, 6\r\nExemplul Sodomei şi Gomorei, 7-8\r\n\r\nd. Metoda învăţătorilor mincinoşi\r\n- pîngăresc trupul, 8\r\n- batjocoresc dregătoriile, 8-10\r\nExemplul lui Cain, 11\r\nExemplul lui Balaam, 11\r\nExemplul fiilor lui Core, 11\r\n\r\ne. Falsitatea învăţătorilor mincinoşi\r\n\r\nŞase metafore care condamnă:\r\n- ca nişte stînci ascunse, 12\r\n- ca nişte nori fără apă, 12\r\n- ca nişte pomi tomnatici fără rod, 12\r\n- ca nişte pomi dezrădăcinaţi, 12\r\n- ca nişte valuri înfuriate ale mării, 13\r\n- ca nişte stele rătăcitoare, 13\r\n\r\nf. Iminenta lor pedepsire\r\nProfeţia patriarhului Enoh, 14-16\r\n\r\nII. LUPTA CELOR CREDINCIOŞI\r\na. A fost vestită de apostoli, 17-19\r\n\r\nb. Trebuie practicată de credincioşi\r\n-” zidiţi-vă sufleteşte\", 20\r\n- „rugaţi-vă prin Duhul\", 20\r\n- „ţineţi-vă în dragostea lui Dumnezeu\", 21\r\n- „aşteptaţi îndurarea Domnului nostru\", 21\r\n- „mustraţi pe cei ce se despart de voi\", 22\r\n- „căutaţi să mîntuiţi pe unii\", 23\r\n- „feriţi-vă de compromisuri cu păcatul\", 23\r\n\r\nIII. NĂDEJDEA CELOR CREDINCIOŞI\r\na. Dumnezeu poate să-i păzească, 24-25" },
    //APOCALIPSA
    { "APOCALIPSA", "Explicatia Cartii: Apocalipsa\r\n\r\nTitlul: În original, cartea poartă numele: „Apokalypsis Iesou Christou\" - „Descoperirea lui Isus Cristos\". Această numire ne atrage atenţia asupra Domnului Isus ca sursă şi subiect general ai tuturor lucrurilor tratate în textul cărţii.\r\n\r\nAutorul: Nu încape nici o îndoială că Cel ce ne trimite această epistolă despre Sine şi despre desfăşurarea istoriei viitoare este însuşi Domnul Isus Cristos. Ioan este numai „instrumentul uman\", „scribul\" ales pentru a ne transmite mesajul primit de la Domnul. Numele lui Ioan apare de patru ori în textul cărţii (Apoc. 1:1, 4, 9; 22:8). Conţinutul cărţii adevereşte şi el că cel ce a scris-o a fost un evreu, cunoscător foarte versat în Scriptură, unul dintre conducătorii spirituali ai Bisericilor din Asia Mică, el însuşi foarte religios şi foarte convins că mişcarea începută de Cristos va triumfa în curînd asupra forţelor demonice care sînt prezente în lume. Apostolul Ioan corespunde cel mai bine unei descrieri ca aceasta.\r\n\r\nData: Cartea Apocalipsei a fost scrisă într-o vreme în care creştinismul intra într-o perioadă de grea persecuţie din partea autorităţilor din Imperiul Roman. Cei mai mulţi comentatori sînt de părere că data scrierii trebuie să fi fost în preajma anului 95 d.Cr.\r\n\r\nContextul scrierii: Din momentul în care autorităţile romane au început să impună în imperiu cultul Cezarului declarat zeu, creştinii - care-L considerau împărat pe Isus şi nu acceptau să i se închine Cezarului - au intrat în conflict deschis cu statul. Apocalipsa îi avertizează pe creştinii din Smirna despre vremurile grele care vor urma (Apoc. 2:10). Antipa, marturul credincios (Apoc. 2:13) căzuse deja împreună cu alţii, ca primele victime produse de persecuţie (Apoc. 6:9). Ioan însuşi se găsea exilat pe insula Patmos (Apoc. l:9), probabil un fel de închisoare a imperiului. Nu este de mirare că, sub presiunile evenimentelor vremii, unii din Biserică începuseră să predice o cale a compromisului (Apoc. 2:14-15, 20), care trebuia combătută repede, mai ales avînd în vedere vremurile şi mai cumplite care trebuiau să vină.\r\n\r\nConţinutul cărţii: Cartea are ca scop să-i încurajeze pe cei credincioşi să stea tari în credinţă şi să nu se plece sub presiunea momentului. Autorul ei îi informează pe cititori că în curînd se va produce confruntarea finală dintre Dumnezeu şi Diavol care se va solda cu zdrobirea Diavolului şi biruinţa glorioasă a Mielului lui Dumnezeu. Pînă atunci însă, creştinii sînt îndemnaţi să stea tare şi să se împotrivească Diavolului chiar şi cu preţul vieţii. Ei trebuie să ştie că au fost pecetluiţi cu sigiliul veşniciei şi că vor fi răzbunaţi la venirea Domnului Isus, cînd cei răi vor fi pedepsiţi pe vecie, iar cei credincioşi vor primi cununa răsplătirilor şi intrarea liberă în eternitatea fericită a unirii desăvîşite cu Fiul lui Dumnezeu.\r\n\r\nPentru a înţelege bine cartea Apocalipsei, cititorul trebuie să ştie că ea este scrisă într-o formă literară specifică. Ezechiel, Daniel, Isaia şi unii dintre profeţi cuprind şi ei pasaje „apocaliptice\". Acest gen de literatură este caracterizat de elemente profund simbolice prin care se încearcă să ni se transmită cunoştinţe despre realităţi care ne depăşesc în mod normal limitele cunoaşterii noastre bazată pe experienţă şi simţuri. Cu toate că la prima vedere viziunile şi imaginile descrise de Ioan par stranii pentru cititorul modern, cartea se poate înţelege deoarece textul însuşi ne pune la dispoziţie „cheia\" unora dintre simbolurile cărţii (de exemplu: stelele sînt îngeri, sfetnicele sînt Biserici - Apoc l:20 - „curva cea mare\" este Babilonul, iar Ierusalimul ceresc este mireasa Mielului - Apoc 21:9-10).\r\n\r\nExistă patru şcoli de interpretare ale cărţii Apocalipsei: cea preteristă, cea idealistă, cea istoricistă şi cea viitoristă. Luîndu-le pe rînd acestea susţin că:\r\n\r\n(1) Preteristă - textul cărţii este simbolic şi legat de evenimentele care au venit asupra Bisericii în secolul I. Astăzi, cartea are doar un caracter documentar, de mărturie a ceea ce s-a întîmplat deja şi din care putem scoate principii veşnic valabile.\r\n\r\n\r\n(2) Idealistă - textul se adresează unor Biserici reale, dar are numai un caracter simbolic, ilustrînd lupta dintre bine şi rău, cu triumful final al binelui. Apocalipsa este redusă la nivelul unei culegeri de fabule.\r\n\r\n(3) Istoricistă - textul este adresat unor Biserici reale, dar capătă un caracter alegoric în care se poate observa o descriere a istoriei din vremea aceea şi pînă la vremea sfîrşitului. Apocalipsa este transformată într-o carte de istorie „cifrată\" care poate rivaliza cu oricare carte de istorie din şcolile lumii. În dosul simbolurilor pot fi recunoscute: căderea Romei, mahomedanismul, papalitatea, reformaţia, etc.\r\n\r\n(4) Viitoristă - natura textului şi felul în care trebuie el tălmăcit sînt reglementate de „cheia\" din Apoc. 1:19: „Scrie lucrurile, pe care le-ai văzut, lucrurile care sînt şi cele care au să vină după ele\". Apocalipsa este privită ca o cronică a vieţii creştine din „vremea Bisericii\" şi ca o anunţare a evenimentelor viitoare din perioada sfîrşitului. Spre deosebire de interpretarea istoricistă, această interpretare nu caută să recunoască istoria trecută a lumii în textul Apocalipsei. Ea rezervă descrierile din carte pentru evenimente încă viitoare, care îşi aşteaptă în curînd împlinirea.\r\n\r\nMajoritatea celor care interpretează „viitorist\" cartea Apocalipsei văd în scrisorile trimise celor şapte Biserici, nu numai nişte epistole cu caracter local, ci descrieri ale unor etape caracteristice prin care va evolua starea creştinismului pînă în vremea sfîrşitului (Apoc. 2 şi 3). Restul cărţii nu s-a întîmplat încă. Evenimentele descrise începînd cu capitolul 4 al cărţii se vor declanşa în preajma sau chiar la cea de a doua venire a Domnului. Capitolul 20 descrie trecerea prin vremea Mileniului spre vremea judecării omenirii, iar ultimele două capitole descriu starea de după judecată, în fericirea eternă a părtăşiei cu Dumnezeu şi cu cerul.\r\n\r\nDeşi recunoaştem ceva bun în fiecare dintre cele patru feluri de interpretare, noi recomandăm tuturor metoda viitoristă. Ea este cel mai aproape de respectarea spiritului Scripturii şi se armonizează cel mai bine cu ceea ce ştim deja din informaţiile transmise nouă grin intermediul celorlalte cărţi profetice. Vechiul Testament a vestit, în repetate ocazii, venirea unei împărăţii în care „Cineva\", venit din linia împărătească a lui David, va domni la Ierusalim peste Israelul refăcut şi reinstalat în propria lui tară, extinzîndu-şi influenţa domniei lui binefăcătoare asupra tuturor neamurilor lumii. Aceste profeţii sînt atît de clare şi în număr atît de mare, că a încerca să le „spiritualizezi\" pe toate, ar însemna o necinstire a inspiraţiei Duhului Sfînt asupra autorilor lor.\r\n\r\nMai există şi acel aspect dublu al lucrării mesianice pe care nu l-au putut înţelege pentru o vreme evreii. Mesia trebuia să vină şi să sufere şi să împărătească. Astăzi, noi ştim că de fapt a fost vorba despre două veniri succesive ale Domnului. Prima dată el a venit în Ierusalim ca să moară pentru păcatele lumii, iar a doua oară se va întoarce în acelaşi Ierusalim ca să-şi instaureze glorioasa Lui împărăţie. Noul Testament nu ne spune cît timp va trebui să treacă între aceste două veniri succesive. El ne dă doar unele evenimente care vor anunţa iminenţa celei de a doua veniri şi declanşarea crizei mondiale care se va sfîrşi cu biruinţa finală a Mielului (Mat. 24:27-31; 2Tes. 2:1-12; 2Tim. 3).\r\n\r\nCuvinte cheie şi teme caracteristice: Istoria lumii se desfăşoară între prima şi ultima carte a Bibliei. Geneza ne arată unde au început toate, iar Apocalipsa ne arată unde şi prin ce se vor sfîrşi toate lucrurile.\r\n\r\nApocalipsa este o succesiune de serii de „şapte\". Se vorbeşte despre: şapte Biserici (Apoc 1:4, 11), şapte duhuri (Apoc 1:4), şapte sfeşnice de aur (Apoc 1:12), şaptestele (Apoc. 1:16), şapte peceţi (Apoc. 5:1), şapte coarne şi şapte ochi (Apoc. 5:6), şapte trîmbiţe (Apoc. 8:2), şapte tunete (Apoc. 10:3), şapte semne (Apoc. 12:1, 3; 13:13-14; 15:1; 16:14; 19:20), şapte cununi împărăteşti (Apoc. 12:3), şapte plăgi (Apoc. 15:7), şapte potire de aur (Apoc. 15:7), şapte munţi (Apoc. 17:9), şi şapte împăraţi (Apoc. 17:10). Cifra „7\" reprezintă în simbolistica iudaică „perfecţiunea divină\". Cartea Apocalipsa ni-L arată pe Dumnezeu la lucru, în toată desăvîrşirea înţelepciunii Lui, contestat, dar nebiruit, atacat, dar mereu la cîrma istoriei, fără să se grăbească şi fără să întîrzie, conducînd totul spre împlinirea planurilor Lui măreţe şi desăvîrşite.\r\n\r\nCartea Apocalipsei poate şi trebuie să fie înţeleasă. Nici o altă carte a Bibliei nu este mai clară în desfăşurarea mesajului ei. Primele cinci capitole descriu prima mişcare a acţiunii prin care Cristos este încununat pe tronul din ceruri. Partea cuprinsă între capitolele 6 şi 20 descrie cea de a doua mişcare a acţiunii spre încununarea lui Cristos ca Domn pe tronul terestru. Finalul cărţii înalţă acţiunea spre apogeul încununării lui Cristos peste toată „noua creaţie\". Cu planul acesta în minte, elementele particulare ale cărţii îşi găsesc repede locul şi semnificaţia.\r\n\r\nCUPRINSUL CĂRŢII\r\n\r\nIntroducere, 1:1-9\r\n\r\nI. CRISTOS - pe tronul cerului\r\na. Fiul Omului între cele şapte sfeşnice, 1:10-20\r\nb. Scrisorile către cele şapte Biserici, 2:1-3:22\r\nc. Tronul slavei şi închinăciunea din cer, 4:1 -5:14\r\n\r\nII. Luarea în stăpînire a pămîntului\r\nd. Ruperea celor şapte peceti, 6:1-17\r\n\r\nO paranteză:\r\nPecetluirea celor 144.000 din Israel, 7:1-8\r\nMulţimea mlhtuitilor din Necazul cel Mare, 7:9-17\r\n\r\ne. Cele şapte trîmbiţe, 8:1-9:21\r\n\r\nO paranteză:\r\nIerusalimul în vremea Necazului cel Mare, 10-11\r\n\r\nf. Cele şapte personaje, 12-13\r\n(femeia însărcinată, pruncul, balaurul roşu, Mihail, vulturul, fiara, a doua fiară)\r\n\r\nO paranteză:\r\nPecetluirea celor 144.000, 14:1-5\r\nVulturul cu Evanghelia veşnică, 14:6-7\r\nAvertismente rmpotriva Minării la fiară, 14:8-13\r\nMotivul pentru Armaghedon, 14:14-20\r\n\r\ng. Cele şapte potire, 15-21\r\n\r\nO paranteză:\r\nBabilonul sub mînia lui Dumnezeu, 17-18\r\n\r\nh. Împărăţia de 1.000 de ani, 20:1-6\r\ni. Ultima împotrivire şi osînda lui Satan, 20:7-10\r\nî. Judecata cea din urmă, 22:11-15\r\n\r\nIII. CRISTOS - domneşte în noua creaţie\r\n\r\nj. Un cer nou şi un pămînt nou, 21-22" }
};


        private readonly List<string> listaCartiBiblice = new List<string>
{
    // Vechiul Testament
    "GENEZA", "EXODUL", "LEVITICUL", "NUMERI", "DEUTERONOMUL",
    "IOSUA", "JUDECĂTORI", "RUT",
    "1 SAMUEL", "2 SAMUEL", "1 IMPARATI", "2 IMPARATI",
    "1 CRONICI", "2 CRONICI", "EZRA", "NEEMIA", "ESTERA",
    "IOV", "PSALMII", "PROVERBE", "ECLESIASTUL", "CÂNTAREA CÂNTĂRILOR",
    "ISAIA", "IEREMIA", "PLANGERILE LUI EREMIA", "EZECHEL", "DANIEL",
    "OSEA", "IOEL", "AMOS", "OBADIA", "IONA",
    "MICA", "NAUM", "HABACUC", "ȚEFANIA", "HAGAI",
    "ZAHARIA", "MALEAHI",

    // Noul Testament
    "MATEI", "MARCU", "LUCA", "IOAN", "FAPTELE APOSTOLILOR",
    "ROMANI", "1 CORINTENI", "2 CORINTENI",
    "GALATENI", "EFESENI", "FILIPENI", "COLOSENI",
    "1 TESALONICENI", "2 TESALONICENI",
    "1 TIMOTEI", "2 TIMOTEI", "TIT", "FILIMON",
    "EVREI", "IACOV",
    "1 PETRU", "2 PETRU",
    "1 IOAN", "2 IOAN", "3 IOAN", "IUDA",
    "APOCALIPSA"
};



        private void LoadCartiCuButoane()
        {
            ListaCarti.Children.Clear();

            foreach (var carte in listaCartiBiblice)
            {
                var panel = new Grid
                {
                    Margin = new Thickness(2),
                    ColumnDefinitions =
            {
                new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) },
                new ColumnDefinition { Width = new GridLength(40) }
            }
                };

                var btnCarte = new Button
                {
                    Content = carte,
                    Tag = carte,
                    FontSize = 14,
                    Padding = new Thickness(6),
                    Margin = new Thickness(2),
                    Height = 40,
                    HorizontalAlignment = HorizontalAlignment.Stretch,
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    MinWidth = 140
                };
                btnCarte.Click += BtnCarte_Click;

                var btnIstoric = new Button
                {
                    Content = "📜",
                    Tag = carte,
                    FontSize = 14,
                    Margin = new Thickness(2),
                    Height = 40,
                    Width = 36,
                    ToolTip = $"Istoric pentru {carte}"   // doar hint scurt
                };
                btnIstoric.Click += BtnIstoric_Click;
                // 🔹 Setăm durata și delay-ul pentru ToolTip
                ToolTipService.SetInitialShowDelay(btnIstoric, 200);   // apare după 0.2 secunde
                ToolTipService.SetShowDuration(btnIstoric, 2500);      // dispare după 1.5 secunde

                Grid.SetColumn(btnCarte, 0);
                Grid.SetColumn(btnIstoric, 1);
                panel.Children.Add(btnCarte);
                panel.Children.Add(btnIstoric);

                ListaCarti.Children.Add(panel);
            }
        }


        private void BtnIstoric_Click(object sender, RoutedEventArgs e)
        {
            var btnIstoric = sender as Button;
            if (btnIstoric == null || btnIstoric.Tag == null)
                return;

            var carte = btnIstoric.Tag.ToString().ToUpper();

            // Dacă popup-ul e deja deschis și e legat de aceeași carte → îl închidem
            if (popup != null && popup.IsOpen && popup.PlacementTarget == btnIstoric)
            {
                popup.IsOpen = false;
                return;
            }

            // Închidem orice popup anterior
            if (popup != null)
                popup.IsOpen = false;

            // Verificăm dacă avem explicație pentru carte
            if (!istoricCarti.ContainsKey(carte))
                return;

            // =========================
            // POPUP (doar la click)
            // =========================

            var tb = ParseInlineWithTags(istoricCarti[carte]); // ✅ un singur TextBlock cu Inlines
            tb.Width = 600;

            var scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                MaxHeight = 600,
                Content = tb
            };
            popup = new Popup
            {
                PlacementTarget = btnIstoric,
                Placement = PlacementMode.Right,
                StaysOpen = true,
                Child = new WpfBorder
                {
                    Background = Brushes.LightYellow,
                    Padding = new Thickness(6),
                    Child = scrollViewer
                }
            };


            popup.IsOpen = true;
        }

        private readonly Dictionary<string, string> mapareCartiNormalizat = new Dictionary<string, string>
        {
            // Vechiul Testament
            ["gen"] = "GENEZA",
            ["exod"] = "EXODUL",
            ["lev"] = "LEVITICUL",
            ["num"] = "NUMERI",
            ["deut"] = "DEUTERONOMUL",
            ["iosua"] = "IOSUA",
            ["jud"] = "JUDECĂTORI",
            ["rut"] = "RUT",
            ["1sam"] = "1 SAMUEL",
            ["2sam"] = "2 SAMUEL",
            ["1 imp"] = "1 IMPARATI",
            ["2 imp"] = "2 IMPARATI",
            ["1cron"] = "1 CRONICI",
            ["2cron"] = "2 CRONICI",
            ["ezra"] = "EZRA",
            ["neem"] = "NEEMIA",
            ["est"] = "ESTERA",
            ["iov"] = "IOV",
            ["ps"] = "PSALMI",
            ["prov"] = "PROVERBE",
            ["ecl"] = "ECLESIASTUL",
            ["cant"] = "CÂNTAREA CÂNTĂRILOR",
            ["isa"] = "ISAIA",
            ["ier"] = "IEREMIA",
            ["plang"] = "PLANGERILE LUI EREMIA",
            ["ezec"] = "EZECHEL",
            ["dan"] = "DANIEL",
            ["osea"] = "OSEA",
            ["ioel"] = "IOEL",
            ["amos"] = "AMOS",
            ["obad"] = "OBADIA",
            ["iona"] = "IONA",
            ["Mica"] = "MICA",
            ["naum"] = "NAUM",
            ["hab"] = "HABACUC",
            ["tef"] = "ȚEFANIA",
            ["hag"] = "HAGAI",
            ["zah"] = "ZAHARIA",
            ["mal"] = "MALEAHI",

            // Noul Testament
            ["mat"] = "MATEI",
            ["marc"] = "MARCU",
            ["luc"] = "LUCA",
            ["ioan"] = "IOAN",
            ["fapte"] = "FAPTELE APOSTOLILOR",
            ["rom"] = "ROMANI",
            ["1corinteni."] = "1 CORINTENI",
            ["2cor"] = "2 CORINTENI",
            ["gal"] = "GALATENI",
            ["efes"] = "EFESENI",
            ["filip"] = "FILIPENI",
            ["col"] = "COLOSENI",
            ["1tes"] = "1 TESALONICENI",
            ["2tes"] = "2 TESALONICENI",
            ["1tim"] = "1 TIMOTEI",
            ["2tim"] = "2 TIMOTEI",
            ["tit"] = "TIT",
            ["filim"] = "FILIMON",
            ["evr"] = "EVREI",
            ["iac"] = "IACOV",
            ["1pet"] = "1 PETRU",
            ["2pet"] = "2 PETRU",
            ["1ioan"] = "1 IOAN",
            ["2ioan"] = "2 IOAN",
            ["3ioan"] = "3 IOAN",
            ["iuda"] = "IUDA",
            ["apoc"] = "APOCALIPSA"
        };


        private string NormalizeCarte(string carte)
        {
            
            carte = RemoveDiacritics(carte)
                .Replace(".", "")
                .Replace(" ", "")
                .ToLowerInvariant();

            return mapareCartiNormalizat.TryGetValue(carte, out var numeComplet)
                ? numeComplet
                : carte.ToUpperInvariant();
        }






        private string RemoveDiacritics(string text)
        {
            var normalized = text.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder();

            foreach (var c in normalized)
            {
                var uc = CharUnicodeInfo.GetUnicodeCategory(c);
                if (uc != UnicodeCategory.NonSpacingMark)
                    sb.Append(c);
            }

            return sb.ToString().Normalize(NormalizationForm.FormC);
        }


        private (string carte, int capitol, string versete) ParseReferinta(string refText)
        {
            // 🔧 Curățare disciplinată
            refText = Regex.Replace(refText, @"[()

\[\]

]", "");                         // elimină paranteze
            refText = refText.TrimEnd('.', ',', ';');                                 // elimină separatori finali
            refText = Regex.Replace(refText, @"\s+", " ").Trim();                     // normalizează spațiile
            refText = Regex.Replace(refText, @"^(?<n>[1-3])(?=\p{L})", "${n} ");      // adaugă spațiu între cifră și literă dacă lipsește

            // 🔍 Regex robust
            var match = Regex.Match(
                refText,
                @"^(?<carte>(?:[1-3]\s?)?\p{L}+\.?(?:\s\p{L}+\.?)*?)\s+(?<capitol>\d+):(?<versete>[\d\-,;]+)$",
                RegexOptions.CultureInvariant
            );

            if (!match.Success)
                throw new FormatException($"Format invalid: {refText}");

            var carte = NormalizeCarte(match.Groups["carte"].Value);
            var capitol = int.Parse(match.Groups["capitol"].Value);
            var versete = match.Groups["versete"].Value;

            return (carte, capitol, versete);
        }









        public TextBlock ParseInlineWithTags(string input)
        {
            var tb = new TextBlock
            {
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 4, 0, 8),
                TextAlignment = System.Windows.TextAlignment.Left,
                FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                FontSize = 16
            };

            string versetPattern =
     @"(?:\(|

\[)?\b(?:Gen\.?|Exod\.?|Lev\.?|Num\.?|Deut\.?|Iosua\.?|Jud\.?|Rut\.?|1\s?Sam\.?|2\s?Sam\.?|1\s?Împ\.?|2\s?Împ\.?|1\s?Cron\.?|2\s?Cron\.?|Ezra\.?|Neem\.?|Est\.?|Iov\.?|Ps\.?|Prov\.?|Ecl\.?|Cant\.?|Isa\.?|Ier\.?|Plang\.?|Ezec\.?|Dan\.?|Osea\.?|Ioel\.?|Amos\.?|Obad\.?|Iona\.?|Mica\.?|Naum\.?|Hab\.?|Țef\.?|Hag\.?|Zah\.?|Mal\.?|Mat\.?|Marc\.?|Luc\.?|Ioan\.?|Fapte\.?|Rom\.?|1\s?Cor\.?|2\s?Cor\.?|Gal\.?|Efes\.?|Filip\.?|Col\.?|1\s?Tes\.?|2\s?Tes\.?|1\s?Tim\.?|2\s?Tim\.?|Tit\.?|Filim\.?|Evr\.?|Iac\.?|1\s?Pet\.?|2\s?Pet\.?|1\s?Ioan\.?|2\s?Ioan\.?|3\s?Ioan\.?|Iuda\.?|Apoc\.?)\s+\d+:\d+(?:-\d+)?(?:\s*(?:[,;]\s*\d+(?:-\d+)?))*\s*(?:[,;])?(?:\)|\]

)?";







            string tagPattern = @"<(?<tag>\w+)>(?<text>.*?)</\k<tag>>";
            int lastIndex = 0;

            // Funcție pentru hyperlink
            void AddHyperlink(string refText)
            {
                refText = Regex.Replace(refText, @"[()\[\].]", ""); // elimină toate parantezele și punctele
                refText = Regex.Replace(refText, @"\s+", " ");      // normalizează spațiile
                refText = refText.TrimEnd(',', ';');                // elimină separator final

                string clean = refText;


                var run = new System.Windows.Documents.Run(refText)
                {
                    FontFamily = new System.Windows.Media.FontFamily("Georgia"),
                    FontSize = 12
                };

                var hyperlink = new System.Windows.Documents.Hyperlink(run)
                {
                    Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 120, 255)),
                    Cursor = Cursors.Hand,
                    TextDecorations = TextDecorations.Underline,
                    ToolTip = GetVersetTooltip(clean)
                };

                hyperlink.Click += async (s, e) =>
                {
                    clean = clean.Trim('(', ')', '[', ']');
                    clean = Regex.Replace(clean, @"\b(\w+)\.", "$1");
                    clean = Regex.Replace(clean, @"\s+", " ");
                    clean = clean.TrimEnd(',', ';');

                    string rezultat;

                    try
                    {
                        var (carte, capitol, versete) = ParseReferinta(clean);
                        rezultat = await Task.Run(() => GetVersetText($"{carte} {capitol}:{versete}"));
                        if (string.IsNullOrEmpty(rezultat))
                        {
                            MessageBox.Show($"Versetul nu a fost găsit: {carte} {capitol}:{versete}");
                            return;
                        }
                    }
                    catch (FormatException ex)
                    {
                        MessageBox.Show(ex.Message);
                        return;
                    }

                    hyperlink.ToolTip = null;

                    var textBlock = new TextBlock
                    {
                        TextWrapping = TextWrapping.Wrap,
                        FontFamily = new System.Windows.Media.FontFamily("Georgia"),
                        FontSize = 18,
                        Foreground = Brushes.Black,
                        TextAlignment = System.Windows.TextAlignment.Left,
                        Margin = new Thickness(20)
                    };

                    string rosupattern = @"<ROSU>(.*?)</ROSU>";
                    var matches = Regex.Matches(rezultat, rosupattern);
                    int li = 0;

                    if (matches.Count == 0)
                    {
                        textBlock.Inlines.Add(new System.Windows.Documents.Run(rezultat));
                    }
                    else
                    {
                        foreach (Match rm in matches)
                        {
                            if (rm.Index > li)
                            {
                                string plain = rezultat.Substring(li, rm.Index - li);
                                textBlock.Inlines.Add(new System.Windows.Documents.Run(plain));
                            }

                            string rosuText = rm.Groups[1].Value;
                            textBlock.Inlines.Add(new System.Windows.Documents.Run(rosuText)
                            {
                                Foreground = Brushes.DarkRed,
                                FontWeight = FontWeights.SemiBold,
                                FontFamily = new System.Windows.Media.FontFamily("Georgia"),
                                FontSize = 16
                            });

                            li = rm.Index + rm.Length;
                        }

                        if (li < rezultat.Length)
                        {
                            string remaining = rezultat.Substring(li);
                            textBlock.Inlines.Add(new System.Windows.Documents.Run(remaining));
                        }
                    }

                    var scrollViewer = new ScrollViewer
                    {
                        Content = textBlock,
                        VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                        HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                        Padding = new Thickness(10)
                    };

                    var win = new Window
                    {
                        Title = $"Trimiterea: {clean}",
                        Content = scrollViewer,
                        Width = 500,
                        Height = 300,
                        WindowStartupLocation = WindowStartupLocation.CenterScreen,
                        Topmost = true,
                        ResizeMode = ResizeMode.CanResizeWithGrip,
                        Background = Brushes.White,
                        BorderBrush = Brushes.Gold,
                        BorderThickness = new Thickness(1)

                    };

                    win.ShowDialog();
                };

                tb.Inlines.Add(hyperlink);
                tb.Inlines.Add(new System.Windows.Documents.Run(" "));
            }

            // Funcție pentru spargere blocuri cu mai multe versete
            void ParseVersetBlock(string block)
            {
                block = block.Trim('(', ')', '[', ']');
                block = Regex.Replace(block, @"\b(\w+)\.", "$1");

                var match = Regex.Match(block, @"^(?<carte>[^\d]+?)\s+(?<capitol>\d+):");
                if (match.Success)
                {
                    string carte = match.Groups["carte"].Value.Trim();
                    string capitol = match.Groups["capitol"].Value;
                    string rest = block.Substring(match.Length).Trim();

                    var versete = new List<string>();
                    if (!string.IsNullOrEmpty(rest))
                    {
                        foreach (var part in rest.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
                        {
                            string vers = part.Trim();
                            if (!vers.Contains(":"))
                                vers = $"{carte} {capitol}:{vers}";
                            versete.Add(vers);
                            vers = Regex.Replace(vers, @"\b(\w+)\.", "$1");
                            vers = vers.Trim('(', ')', '[', ']');

                        }
                    }
                    else
                    {
                        versete.Add($"{carte} {capitol}:{block.Substring(match.Length).Trim()}");

                    }

                    foreach (var v in versete)
                        AddHyperlink(v);
                }
                else
                {
                    AddHyperlink(block);
                }
            }

            // Procesare taguri
            foreach (Match m in Regex.Matches(input, tagPattern, RegexOptions.Singleline))
            {
                if (m.Index > lastIndex)
                {
                    string plainText = input.Substring(lastIndex, m.Index - lastIndex);
                    tb.Inlines.Add(new System.Windows.Documents.Run(plainText));
                }

                string tag = m.Groups["tag"].Value;
                string text = m.Groups["text"].Value;

                
                    switch (tag)
                {
                    case "titlu":
                        tb.Inlines.Add(new System.Windows.Documents.Run(text)
                        {
                            FontFamily = new System.Windows.Media.FontFamily("Georgia"),
                            FontSize = 22,
                            FontWeight = FontWeights.Bold,
                            Foreground = System.Windows.Media.Brushes.DarkRed
                        });
                        break;

                    case "autor":
                        tb.Inlines.Add(new System.Windows.Documents.Run(text)
                        {
                            FontFamily = new System.Windows.Media.FontFamily("Georgia"),
                            FontSize = 20,
                            FontStyle = FontStyles.Italic,
                            Foreground = System.Windows.Media.Brushes.DarkBlue
                        });
                        break;

                    case "explicație":
                        tb.Inlines.Add(new System.Windows.Documents.Run(text)
                        {
                            FontFamily = new System.Windows.Media.FontFamily("Georgia"),
                            FontSize = 18,
                            Foreground = System.Windows.Media.Brushes.DarkGreen
                        });
                        break;

                    case "trimitere":
                        var fragmente = Regex.Split(text, versetPattern);
                        var versete = Regex.Matches(text, versetPattern);
                        for (int i = 0; i < fragmente.Length; i++)
                        {
                            tb.Inlines.Add(new System.Windows.Documents.Run(fragmente[i]));
                            if (i < versete.Count)
                                ParseVersetBlock(versete[i].Value);
                        }
                        break;

                    case "ROSU":
                        tb.Inlines.Add(new System.Windows.Documents.Run(text)
                        {
                            Foreground = System.Windows.Media.Brushes.DarkRed,
                            FontWeight = FontWeights.SemiBold,
                            FontFamily = new System.Windows.Media.FontFamily("Georgia"),
                            FontSize = 16
                        });
                        break;

                    default:
                        tb.Inlines.Add(new System.Windows.Documents.Run(text));
                        break;
                }



                lastIndex = m.Index + m.Length;
            }

            // Text rămas în afara tagurilor
            if (lastIndex < input.Length)
            {
                string remainingText = input.Substring(lastIndex);
                int lastPos = 0;
                foreach (Match m in Regex.Matches(remainingText, versetPattern))
                {
                    if (m.Index > lastPos)
                        tb.Inlines.Add(new System.Windows.Documents.Run(remainingText.Substring(lastPos, m.Index - lastPos)));

                    ParseVersetBlock(m.Value);
                    lastPos = m.Index + m.Length;
                }

                if (lastPos < remainingText.Length)
                    tb.Inlines.Add(new System.Windows.Documents.Run(remainingText.Substring(lastPos)));
            }

            return tb;
        }







        private string GetVersetText(string referinta)
        {
            try
            {
                // 🔹 Parsează referința
                var (carte, capitol, verseteRaw) = ParseReferinta(referinta);

                // 🔹 Diagnostic silențios — scrie în fișier
                File.AppendAllText("diagnostic.txt", $"Carte={carte}, Capitol={capitol}, Versete={verseteRaw}\n");

                // 🔹 Parsează lista de versete
                var verseteList = new List<int>();
                foreach (var part in verseteRaw.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (part.Contains("-"))
                    {
                        var range = part.Split('-');
                        int start = int.Parse(range[0]);
                        int end = int.Parse(range[1]);
                        for (int v = start; v <= end; v++)
                            verseteList.Add(v);
                    }
                    else
                    {
                        verseteList.Add(int.Parse(part));
                    }
                }

                // 🔹 Interogare DB
                var rezultate = new List<string>();
                using (var conn = new SQLiteConnection("Data Source=biblia.db"))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = @"
                    SELECT Verset, Text FROM Versete
                    WHERE Carte = @carte AND Capitol = @capitol AND Verset = @verset
                    ORDER BY Verset";

                        cmd.Parameters.AddWithValue("@carte", carte);
                        cmd.Parameters.AddWithValue("@capitol", capitol);
                        cmd.Parameters.Add("@verset", System.Data.DbType.Int32);

                        foreach (int v in verseteList)
                        {
                            cmd.Parameters["@verset"].Value = v;
                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    int nr = reader.GetInt32(0);
                                    string text = reader.GetString(1);
                                    rezultate.Add($"{nr}. {text}");
                                }
                            }
                        }
                    }
                }

                return rezultate.Count > 0
                    ? string.Join("\n", rezultate)
                    : $"[Versetul nu a fost găsit: {carte} {capitol}:{verseteRaw}]";
            }
            catch (Exception ex)
            {
                File.AppendAllText("diagnostic.txt", $"EROARE: {ex.Message}\n");
                return $"[Eroare la căutare: {ex.Message}]";
            }
        }











        private ToolTip GetVersetTooltip(string referinta)
        {
            var tb = new TextBlock
            {
                Text = $"Trimiterea BIBLICĂ: {referinta}",
                Foreground = Brushes.DarkBlue,
                FontWeight = FontWeights.Bold,
                FontSize = 12
            };

            return new ToolTip { Content = tb };
        }













        private void BtnCarte_MouseEnter(object sender, MouseEventArgs e)
        {
            var btnCarte = sender as Button;
            var carte = btnCarte.Tag.ToString();

            if (istoricCarti.ContainsKey(carte.ToUpper()))
            {
                var textBlock = new TextBlock
                {
                    Text = istoricCarti[carte.ToUpper()],
                    Foreground = Brushes.DarkBlue,
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    TextWrapping = TextWrapping.Wrap,
                    Width = 550
                };

                var scrollViewer = new ScrollViewer
                {
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    MaxHeight = 600,
                    Content = textBlock
                };

                // 🔹 Închidem popup-ul anterior dacă există
                if (popup != null)
                    popup.IsOpen = false;

                popup = new Popup
                {
                    PlacementTarget = btnCarte,
                    Placement = PlacementMode.Right,   // 🔹 apare în dreapta butonului
                    HorizontalOffset = 0,              // 🔹 ajustează stânga/dreapta
                    VerticalOffset = 0,                // 🔹 ajustează sus/jos
                    StaysOpen = true,
                    Child = new System.Windows.Controls.Border
                    {
                        Background = Brushes.LightYellow,
                        Padding = new Thickness(6),
                        Child = scrollViewer
                    }
                };


                popup.IsOpen = true;
            }
        }

        private void BtnCarte_Click(object sender, RoutedEventArgs e)
        {
            // 🔹 Închidem popup-ul dacă e deschis
            if (popup != null)
                popup.IsOpen = false;

            var btnCarte = sender as Button;
            var carte = btnCarte.Tag.ToString();

            // aici vine logica ta de încărcare capitole
            // ex:
            ListaCapitole.Children.Clear();
            ScrollCapitole.Visibility = Visibility.Visible;

            var cmd = new SQLiteCommand("SELECT DISTINCT Capitol FROM Versete WHERE Carte=@carte ORDER BY Capitol", conn);
            cmd.Parameters.AddWithValue("@carte", carte);
            var reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                var capitol = reader["Capitol"].ToString();
                var btnCapitol = new Button
                {
                    Content = capitol,
                    Tag = System.Tuple.Create(carte, capitol),
                    Margin = new Thickness(3),
                    Padding = new Thickness(10),
                   
                    FontSize = 14
                };

                btnCapitol.Click += BtnCapitol_Click;
                ListaCapitole.Children.Add(btnCapitol);
            }
        }



        private void BtnCapitol_Click(object sender, RoutedEventArgs e)
        {
            if (popup != null)
                popup.IsOpen = false;

            if (sender is not Button btnCapitol) return;

            var (carte, capitol) = (Tuple<string, string>)btnCapitol.Tag;

            versetCarteCurenta = carte;
            versetCapitolCurent = int.Parse(capitol);

            lblPozitieCurenta.Content = $"📖 {carte} {capitol}";

            ActualizeazaTitlu();
            IncarcaVersete(versetCarteCurenta, versetCapitolCurent);


        }





        private void btnToggleTrimiteri_Click(object sender, RoutedEventArgs e)
        {
            trimiteriVizibile = !trimiteriVizibile;

            // ToolTip stilizat în funcție de stare
            btnToggleTrimiteri.ToolTip = new ToolTip
            {
                Background = Brushes.LightYellow,
                Padding = new Thickness(6),
                Content = new TextBlock
                {
                    Text = trimiteriVizibile ? "Ascunde trimiterile" : "Arată trimiterile",
                    Foreground = trimiteriVizibile ? Brushes.Green : Brushes.DarkBlue,
                    FontSize = 12,
                    FontWeight = FontWeights.Bold
                }
            };

            // timpii pentru afișare instant
            ToolTipService.SetInitialShowDelay(btnToggleTrimiteri, 0);
            ToolTipService.SetShowDuration(btnToggleTrimiteri, 5000);
            ToolTipService.SetBetweenShowDelay(btnToggleTrimiteri, 0);

            if (versetCarteCurenta != null && versetCapitolCurent > 0)
            {
                IncarcaVersete(versetCarteCurenta, versetCapitolCurent);
            }
        }


        public void IncarcaVersete(string carte, int capitol)
        {

            SeteazaMod(ModAfisare.Scriptura);

            listVersete.Items.Clear();

            var cmd = new SQLiteCommand(@"
        SELECT Verset, Text, Trimiteri 
        FROM Versete 
        WHERE Carte = @carte AND Capitol = @capitol 
        ORDER BY Verset", conn);

            cmd.Parameters.AddWithValue("@carte", carte);
            cmd.Parameters.AddWithValue("@capitol", capitol);

            var reader = cmd.ExecuteReader();

            if (!reader.HasRows)
            {
                MessageBox.Show($"Nu s-au găsit versete pentru {carte} {capitol}");
                return;
            }

            while (reader.Read())
            {
                var verset = reader.GetInt32(0);
                var text = reader.GetString(1);
              

                var textTrimiteri = reader.IsDBNull(2) ? "" : reader.GetString(2);

                var wrap = new WrapPanel
                {
                    Orientation = Orientation.Horizontal,
                    MaxWidth = listVersete.ActualWidth - 20,
                    Margin = new Thickness(0, 1, 0, 1),
                    VerticalAlignment = VerticalAlignment.Top
                };

                var tb = new TextBlock
                {
                    TextWrapping = TextWrapping.Wrap,
                    FontSize = marimeTextCurenta,
                    FontFamily = new WpfFontFamily("Segoe UI"),
                    Foreground = culoareTextVerset
                };
                tb.Inlines.Add(new WpfRun($"{verset}. ")
                {
                    Foreground = culoareTextVerset
                });

                // AICI PROCESEZI TAGURILE <ROSU>
                if (!string.IsNullOrWhiteSpace(cuvantCurentCautat))
                {
                    foreach (var inline in ParseVersetCuRosu(text, cuvantCurentCautat))
                        tb.Inlines.Add(inline);
                }
                else
                {
                    foreach (var inline in ParseTextCuTagRosu(text))
                        tb.Inlines.Add(inline);
                }







                // 🔹 Trimiteri biblice
                if (trimiteriVizibile && !string.IsNullOrWhiteSpace(textTrimiteri))
                {
                   
                    

                    var regex = new Regex(@"\b([1-3]?\s?[A-ZĂÂÎȘȚa-zăâîșț\.]+)\s+(\d+):(\d+)\b");
                    var matches = regex.Matches(textTrimiteri);

                    
                    tb.Inlines.Add(new System.Windows.Documents.Run("")
                    {
                        FontStyle = FontStyles.Italic,
                        Foreground = Brushes.DarkRed
                    });


                    foreach (Match match in matches)
                    {
                        string carteTrim = match.Groups[1].Value.Trim().Replace(".", "");
                        int capitolTrim = int.Parse(match.Groups[2].Value);
                        int versetTrim = int.Parse(match.Groups[3].Value);

                        var hyperlink = new System.Windows.Documents.Hyperlink(
                            new System.Windows.Documents.Run($"{carteTrim} {capitolTrim}:{versetTrim}")
                            {
                                FontSize = marimeTextCurenta -6,
                                FontWeight = FontWeights.Bold
                            }
                        )
                        {
                            Foreground = Brushes.SteelBlue,
                            Cursor = Cursors.Hand,
                            FontSize = marimeTextCurenta + 2
                        };

                        hyperlink.Click += (s, e) =>
                        {
                            string carteTrimNorm = NormalizareCarte(carteTrim.Replace(".", "").Trim());
                            var fereastraNoua = new FereastraTrimiteri(
                                this,
                                carteTrimNorm,
                                capitolTrim,
                                new List<int> { versetTrim }
                            );

                            fereastraNoua.Show();
                        };

                        tb.Inlines.Add(new System.Windows.Documents.Run(" "));
                        tb.Inlines.Add(hyperlink);
                    }

                }

                wrap.Children.Add(tb);
                listVersete.Items.Add(new ListBoxItem
                {
                    Content = wrap,
                    Foreground = culoareTextVerset
                });
                // 🔹 Buton 📘 Comentariu
                if (modComentariiActiv && AreComentariu(carte, capitol, verset))
                {
                    var btnComentariu = new Button
                    {
                        Content = "📘",
                        FontSize = 12,
                        Width = 22,
                        Height = 22,
                        Margin = new Thickness(6, 0, 0, 0),
                        VerticalAlignment = VerticalAlignment.Center,
                        Tag = new VersetSchita
                        {
                            Carte = carte,
                            Capitol = capitol,
                            Verset = verset,
                            Text = text
                        }
                    };

                    // 👉 ToolTip stilizat
                    btnComentariu.ToolTip = new ToolTip
                    {
                        Background = Brushes.LightYellow,   // fundal cald
                        Padding = new Thickness(6),         // spațiu intern
                        Content = new TextBlock
                        {
                            Text = "Vezi comentariul 📘",
                            Foreground = Brushes.DarkBlue, // text verde închis
                            FontSize = 14,
                            FontWeight = FontWeights.Bold
                        }
                    };

                    // timpii pentru afișare instant
                    ToolTipService.SetInitialShowDelay(btnComentariu, 10);
                    ToolTipService.SetShowDuration(btnComentariu, 5000);
                    ToolTipService.SetBetweenShowDelay(btnComentariu, 0);

                    btnComentariu.Click += BtnComentariu_Click;
                    wrap.Children.Add(btnComentariu);


                }

                // 🔹 Buton ⭐ Favorite
                if (arataStele)
                {
                    var btnFavorite = new Button
                    {
                        Content = "⭐",
                        FontSize = 12,
                        Width = 22,
                        Height = 22,
                        Margin = new Thickness(6, 0, 0, 0),
                        VerticalAlignment = VerticalAlignment.Center,
                        ToolTip = "Adaugă la favorite",
                        Tag = new VersetSchita { Carte = carte, Capitol = capitol, Verset = verset, Text = text }
                    };
                    btnFavorite.Click += BtnAdaugaFavorite_Click;
                    wrap.Children.Add(btnFavorite);
                }


                // 🔹 Buton ➕ / ✅ Schiță
                if (modSchitaActiv)
                {
                    bool esteInSchita = schitaCurenta.Any(v =>
                        v.Carte == carte &&
                        v.Capitol == capitol &&
                        v.Verset == verset);

                    var btnAdauga = new Button
                    {
                        Content = esteInSchita ? "✅" : "➕",
                        FontSize = 12,              // mai mic decât 16
                        Width = 22,                 // micșorează lățimea
                        Height = 22,                // micșorează înălțimea
                        Padding = new Thickness(0), // elimină spațiul intern
                        Margin = new Thickness(4, 0, 0, 0),
                        VerticalAlignment = VerticalAlignment.Center,
                        ToolTip = esteInSchita ? "Versetul este deja în schiță" : "Adaugă la schiță",
                        IsEnabled = !esteInSchita,
                        Tag = new VersetSchita
                        {
                            Carte = carte,
                            Capitol = capitol,
                            Verset = verset,
                            Text = text
                        }
                    };

                    // 👉 aici setezi timpii pentru afișare instant
                    ToolTipService.SetInitialShowDelay(btnAdauga, 10);   // apare imediat la hover
                    ToolTipService.SetShowDuration(btnAdauga, 5000);    // rămâne vizibil 5 secunde
                    ToolTipService.SetBetweenShowDelay(btnAdauga, 0);   // reapare imediat dacă muți cursorul și revii

                    if (!esteInSchita)
                        btnAdauga.Click += BtnAdaugaLaSchita_Click;

                    btnAdauga.ToolTip = new ToolTip
                    {
                        Background = Brushes.LightYellow,   // fundalul balonului
                        Padding = new Thickness(6),         // spațiu intern
                        Content = new TextBlock
                        {
                            Text = esteInSchita ? "Versetul este deja în schiță" : "Adaugă la schiță",
                            Foreground = Brushes.DarkBlue,  // culoarea textului
                            FontSize = 18,                  // dimensiunea fontului
                            FontWeight = FontWeights.Bold   // text îngroșat
                        }
                    };



                    wrap.Children.Add(btnAdauga);
                }

                listVersete.Items.Add(new ListBoxItem
                {
                    Content = wrap,
                    Foreground = culoareTextVerset
                });
            }

            if (listVersete.Items.Count > 0)
            {
                listVersete.SelectedIndex = 0;
                listVersete.ScrollIntoView(listVersete.SelectedItem);
            }
        }
        private string NormalizareCarte(string prescurtare)
        {
            
            prescurtare = prescurtare.Trim().TrimStart('*', '†', '‡', '✝').TrimEnd('.').ToUpper();
            

            switch (prescurtare)
            {
                case "GEN": return "GENEZA";
                case "EXOD": return "EXODUL";
                case "LEV": return "LEVITICUL";
                case "NUM": return "NUMERI";
                case "DEUT": return "DEUTERONOMUL";
                case "IOS": return "IOSUA";
                case "JUD": return "JUDECĂTORI";
                case "RUT": return "RUT";
                case "1SAM": return "1 SAMUEL";
                case "1 SAM": return "1 SAMUEL";
                case "2SAM": return "2 SAMUEL";
                case "2 SAM": return "1 SAMUEL";
                case "1IMP": return "1 IMPARATI";
                case "1 IMP": return "1 IMPARATI";
                case "2IMP": return "2 IMPARATI";
                case "2 IMP": return "1 IMPARATI";
                case "1CRON": return "1 CRONICI";
                case "1 CRON": return "1 CRONICI";
                case "2CRON": return "2 CRONICI";
                case "2 CRON": return "2 CRONICI";
                case "EZRA": return "EZRA";
                case "NEEM": return "NEEMIA";
                case "EST": return "ESTERA";
                case "IOV": return "IOV";
                case "PS": return "PSALMII";
                case "PROV": return "PROVERBE";
                case "ECCL": return "ECLEZIASTUL";
                case "CANT": return "CÂNTAREA CÂNTĂRILOR";
                case "ISA": return "ISAIA";
                case "IER": return "IEREMIA";
                case "PLNG": return "PLÂNGERILE";
                case "EZEC": return "EZECHEL";
                case "DAN": return "DANIEL";
                case "OSEA": return "OSEA";
                case "IOEL": return "IOEL";
                case "AMOS": return "AMOS";
                case "OBAD": return "OBADIA";
                case "IONA": return "IONA";
                case "MIH": return "MICA"; // corectat
                case "NAUM": return "NAUM";
                case "HAB": return "HABACUC";
                case "TCEF": return "ȚEFANIA";
                case "HAG": return "HAGAI";
                case "ZAH": return "ZAHARIA";
                case "MAL": return "MALEAHI";
                case "MAT": return "MATEI";
                case "MAR": return "MARCU";
                case "LUC": return "LUCA";
                case "IOAN": return "IOAN";
                case "FAPTE": return "FAPTELE APOSTOLILOR";
                case "ROM": return "ROMANI";
                case "1COR": return "1 CORINTENI";
                case "2COR": return "2 CORINTENI";
                case "GAL": return "GALATENI";
                case "EFES": return "EFESENI";
                case "FIL": return "FILIPENI";
                case "COL": return "COLOSENI";
                case "1TES": return "1 TESALONICENI";
                case "2TES": return "2 TESALONICENI";
                case "1TIM": return "1 TIMOTEI";
                case "2TIM": return "2 TIMOTEI";
                case "TIT": return "TIT";
                case "FILIM": return "FILIMON";
                case "EVR": return "EVREI";
                case "IAC": return "IACOV";
                case "1PET": return "1 PETRU";
                case "2PET": return "2 PETRU";
                case "1IOAN": return "1 IOAN";
                case "2IOAN": return "2 IOAN";
                case "3IOAN": return "3 IOAN";
                case "IUDA": return "IUDA";
                case "APO": return "APOCALIPSA";
                default: return prescurtare;
            }

        }



        private static string ExtrageLocatie(string text)
        {
            if (string.IsNullOrEmpty(text)) return null;

            // elimină paranteze și spații inutile
            text = text.Replace("(", " ").Replace(")", " ").Trim();

            // caută modelul capitol:verset(-verset)
            var m = Regex.Match(text, @"\d{1,3}:\d{1,3}(?:-\d{1,3})?");
            if (m.Success) return m.Value;

            return null;
        }

        private int _clickCount = 0;
        private static string NormalizeRefText(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return s;
            s = Regex.Replace(s, @"\u2013|\u2014", "-");      // –, — → -
            s = Regex.Replace(s, @"\s+", " ").Trim();
            s = Regex.Replace(s, @"([A-Za-zĂÂÎȘȚăâîșț]+)\.\s*(\d{1,3}:\d{1,3})", "$1 $2"); // Marcu. 1:21 → Marcu 1:21
            return s;
        }
        private static string PrimaMajuscula(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;
            text = text.ToLowerInvariant();
            return char.ToUpper(text[0]) + text.Substring(1);
        }

        private void Hyperlink_MouseLeftButtonUp(object sender, RoutedEventArgs e)

        {
            e.Handled = true;

            if (sender is WpfHyperlink link)
            {
               

                if (link?.NavigateUri == null)
                {
                    MessageBox.Show("NavigateUri este null!");
                    return;
                }

                string encoded = link.NavigateUri.ToString()
                    .Replace("internal://", "")
                    .Replace("_", " ")
                    .Trim();
                

                string locatie = NormalizeLocatie(ExtrageLocatie(encoded));
                if (string.IsNullOrEmpty(locatie))
                {
                    MessageBox.Show("Locația nu a putut fi extrasă!");
                    return;
                }

                string[] parti = locatie.Split(':');
                if (parti.Length < 2)
                {
                    MessageBox.Show("Locația nu are format capitol:verset!");
                    return;
                }

                string carte = NormalizareCarte(encoded.Split(' ')[0]);
                if (string.IsNullOrEmpty(carte))
                {
                    MessageBox.Show("Cartea nu a putut fi normalizată!");
                    return;
                }

                int capitol = int.Parse(parti[0]);
                int verset = int.Parse(parti[1]);

                var fereastra = new FereastraTrimiteri(
      this,              // BibliaWindow (fereastra principală)
      carte,             // string
      capitol,           // int
      new List<int> { verset } // lista de versete
  );

                fereastra.Owner = this;
                fereastra.Show();
            }
        }














        private static string NormalizeLocatie(string loc)
        {
            if (string.IsNullOrEmpty(loc)) return loc;

            return loc.Replace('\u00A0', ' ')
                      .Replace('–', '-')
                      .Replace('—', '-')
                      .Trim('(', ')', '.', ':', ';', ' ');
        }



        private void btnToggleComentarii_Click(object sender, RoutedEventArgs e)
        {
            modComentariiActiv = !modComentariiActiv;

            // ToolTip stilizat în funcție de stare
            btnToggleComentarii.ToolTip = new ToolTip
            {
                Background = Brushes.LightYellow,
                Padding = new Thickness(6),
                Content = new TextBlock
                {
                    Text = modComentariiActiv ? "Ascunde comentariile 📘" : "Arată comentariile 📘",
                    Foreground = modComentariiActiv ? Brushes.Green : Brushes.DarkBlue,
                    FontSize = 12,
                    FontWeight = FontWeights.Bold
                }
            };

            // timpii pentru afișare instant
            ToolTipService.SetInitialShowDelay(btnToggleComentarii, 0);
            ToolTipService.SetShowDuration(btnToggleComentarii, 5000);
            ToolTipService.SetBetweenShowDelay(btnToggleComentarii, 0);

            if (!string.IsNullOrEmpty(versetCarteCurenta) && versetCapitolCurent > 0)
                IncarcaVersete(versetCarteCurenta, versetCapitolCurent);

            ActualizeazaTitlu();
        }









        private List<Inline> ParseVersetCuRosu(string text, string cuvant)
        {
            var inlines = new List<Inline>();
            var regexRosu = new Regex(@"<\s*ROSU\s*>([\s\S]*?)<\s*/\s*ROSU\s*>", RegexOptions.IgnoreCase);

            int lastIndex = 0;

            foreach (Match match in regexRosu.Matches(text))
            {
                // textul dinainte
                if (match.Index > lastIndex)
                {
                    string normal = text.Substring(lastIndex, match.Index - lastIndex);
                    inlines.AddRange(EvidentiazaCuvant(normal, cuvant));
                }

                // textul dintre taguri <ROSU> — evidențiat și colorat
                string rosu = match.Groups[1].Value;
                foreach (var inline in EvidentiazaCuvant(rosu, cuvant))
                {
                    if (inline is WpfRun run)
                    {
                        run.Foreground = Brushes.Red;
                        run.FontWeight = FontWeights.Bold;
                    }

                    inlines.Add(inline);  // ✔ AICI adaugi inline, nu run
                }


                lastIndex = match.Index + match.Length;
            }

            // textul de după ultimul tag
            if (lastIndex < text.Length)
            {
                string final = text.Substring(lastIndex);
                inlines.AddRange(EvidentiazaCuvant(final, cuvant));
            }

            return inlines;
        }



        private List<Inline> EvidentiazaCuvant(string text, string cuvant)
        {
            var inlines = new List<Inline>();
            cuvant = cuvant.ToLower();

            // Procesăm textul cu taguri <rosu>
            int index = 0;
            while (index < text.Length)
            {
                int startRosu = text.IndexOf("<rosu>", index);
                if (startRosu == -1)
                {
                    // Nu mai sunt taguri <rosu> — procesăm restul normal
                    string rest = text.Substring(index);
                    inlines.AddRange(EvidentiazaCuvantSimplu(rest, cuvant, Brushes.Black));
                    break;
                }

                // Text normal înainte de <rosu>
                string normal = text.Substring(index, startRosu - index);
                inlines.AddRange(EvidentiazaCuvantSimplu(normal, cuvant, Brushes.Black));

                // Text între <rosu> și </rosu>
                int endRosu = text.IndexOf("</rosu>", startRosu);
                if (endRosu == -1)
                {
                    // Tag închis greșit — tratăm restul ca roșu
                    string rosu = text.Substring(startRosu + 6);
                    inlines.AddRange(EvidentiazaCuvantSimplu(rosu, cuvant, Brushes.Red));
                    break;
                }

                string rosuText = text.Substring(startRosu + 6, endRosu - (startRosu + 6));
                inlines.AddRange(EvidentiazaCuvantSimplu(rosuText, cuvant, Brushes.Red));

                index = endRosu + 7;
            }

            return inlines;
        }
        private List<Inline> EvidentiazaCuvantSimplu(string text, string cuvant, Brush culoare)
        {
            var inlines = new List<Inline>();
            int pos = 0;
            string textLower = text.ToLower();

            while (pos < text.Length)
            {
                int index = textLower.IndexOf(cuvant, pos);
                if (index == -1)
                {
                    inlines.Add(new Run(text.Substring(pos)) { Foreground = culoare });
                    break;
                }

                // Text înainte de cuvânt
                if (index > pos)
                    inlines.Add(new Run(text.Substring(pos, index - pos)) { Foreground = culoare });

                // Cuvântul evidențiat
                inlines.Add(new Run(text.Substring(index, cuvant.Length))
                {
                    Foreground = Brushes.Goldenrod,
                    FontWeight = FontWeights.Bold
                });

                pos = index + cuvant.Length;
            }

            return inlines;
        }







        private string DiacriticeTolerante(string cuvant)
        {
            return cuvant
                .Replace("a", "[aăâ]")
                .Replace("A", "[AĂÂ]")
                .Replace("i", "[iî]")
                .Replace("I", "[IÎ]")
                .Replace("s", "[sșşşș]")
                .Replace("S", "[SȘŞŞȘ]")
                .Replace("t", "[tțţţț]")
                .Replace("T", "[TȚŢŢȚ]");
        }





        private string Normalize(string s)
        {
            string rezultat = s.ToLowerInvariant();

            // Diacritice
            rezultat = rezultat.Replace("ă", "a").Replace("â", "a").Replace("î", "i")
                               .Replace("ș", "s").Replace("ş", "s").Replace("ț", "t").Replace("ţ", "t")
                               .Replace("ş", "s").Replace("ţ", "t").Replace("ș", "s").Replace("ț", "t");

            // Majuscule diacritice (dacă apar)
            rezultat = rezultat.Replace("Ă", "a").Replace("Â", "a").Replace("Î", "i")
                               .Replace("Ș", "s").Replace("Ț", "t");

            // Semne de punctuație și caractere speciale
            string[] semne = { ".", ",", ";", ":", "-", "—", "*", "†", "„", "”", "!", "?", "(", ")", "[", "]", "\"" };
            foreach (var semn in semne)
                rezultat = rezultat.Replace(semn, " ");

            // Spații duble (de 3 ori)
            for (int i = 0; i < 3; i++)
                rezultat = rezultat.Replace("  ", " ");

            // Spațiu la început și sfârșit
            rezultat = " " + rezultat.Trim() + " ";

            return rezultat;
        }



        private int FindRealIndex(string original, string normalized, int targetNormIndex)
        {
            int realIndex = 0;
            int normIndex = 0;

            while (realIndex < original.Length && normIndex < targetNormIndex)
            {
                string c = original[realIndex].ToString();
                normIndex += Normalize(c).Length;
                realIndex++;
            }

            return realIndex;
        }


        


        private void Evidentiaza(TextBlock tb, string textOriginal, string cuvant)
        {
            tb.Inlines.Clear();

            if (string.IsNullOrWhiteSpace(cuvant))
            {
                tb.Text = textOriginal;
                return;
            }

            string pattern = $@"\b{Regex.Escape(cuvant)}\b";

            var parts = Regex.Split(textOriginal, pattern, RegexOptions.IgnoreCase);

            int index = 0;

            foreach (var part in parts)
            {
                // text normal
                tb.Inlines.Add(new System.Windows.Documents.Run(part));

                // dacă mai există o potrivire după acest segment
                index += part.Length;
                if (index < textOriginal.Length)
                {
                    var match = Regex.Match(textOriginal.Substring(index), pattern, RegexOptions.IgnoreCase);
                    if (match.Success && match.Index == 0)
                    {
                        tb.Inlines.Add(new System.Windows.Documents.Run(match.Value)
                        {
                            Foreground = Brushes.Red,
                            FontWeight = FontWeights.Bold
                        });
                        index += match.Length;
                    }
                }
            }
        }










        private double marimeTextCurenta = 17;

        private void btnMarimeText_Click(object sender, RoutedEventArgs e)
        {
            marimeTextCurenta += 2;
            if (marimeTextCurenta > 30) marimeTextCurenta = 14; // revine la mic

            // ToolTip stilizat în funcție de mărime
            btnMarimeText.ToolTip = new ToolTip
            {
                Background = Brushes.LightYellow,
                Padding = new Thickness(6),
                Content = new TextBlock
                {
                    Text = $"Dimensiunea textului: {marimeTextCurenta}",
                    Foreground = Brushes.DarkBlue,
                    FontSize = 14,
                    FontWeight = FontWeights.Bold
                }
            };

            ToolTipService.SetInitialShowDelay(btnMarimeText, 0);
            ToolTipService.SetShowDuration(btnMarimeText, 5000);
            ToolTipService.SetBetweenShowDelay(btnMarimeText, 0);

            if (versetCarteCurenta != null && versetCapitolCurent > 0)
            {
                IncarcaVersete(versetCarteCurenta, versetCapitolCurent);
            }
        }


        private bool modIntunecatActiv = false;

        private void btnModIntunecat_Click(object sender, RoutedEventArgs e)
        {
            modIntunecatActiv = !modIntunecatActiv;

            // Fundal principal
            var fundal = modIntunecatActiv ? Brushes.Black : Brushes.White;
            Background = fundal;

            // Meniu lateral
            MeniuLateral.Background = modIntunecatActiv ? Brushes.DarkSlateGray : Brushes.LightGray;

            // Texturi și etichete
            lblPozitieCurenta.Foreground = culoareTextVerset;

            // ToolTip stilizat în funcție de stare
            btnModIntunecat.ToolTip = new ToolTip
            {
                Background = Brushes.LightYellow,
                Padding = new Thickness(6),
                Content = new TextBlock
                {
                    Text = modIntunecatActiv ? "Mod luminos" : "Mod întunecat",
                    Foreground = modIntunecatActiv ? Brushes.Green : Brushes.DarkBlue,
                    FontSize = 14,
                    FontWeight = FontWeights.Bold
                }
            };

            ToolTipService.SetInitialShowDelay(btnModIntunecat, 0);
            ToolTipService.SetShowDuration(btnModIntunecat, 5000);
            ToolTipService.SetBetweenShowDelay(btnModIntunecat, 0);

            // Fundal și text pentru lista de versete
            listVersete.Background = fundal;
            listVersete.Foreground = culoareTextVerset;

            // Fundal și text pentru căutare
            txtCautare.Background = modIntunecatActiv ? Brushes.DarkSlateGray : Brushes.White;
            txtCautare.Foreground = culoareTextVerset;

            // Fundal pentru fiecare item deja încărcat (dacă există)
            foreach (var item in listVersete.Items)
            {
                if (item is ListBoxItem lbi && lbi.Content is TextBlock tb)
                {
                    lbi.Background = fundal;
                    lbi.Foreground = culoareTextVerset;
                    tb.Foreground = culoareTextVerset;

                    foreach (var inline in tb.Inlines)
                    {
                        if (inline is WpfRun run && run.Background != Brushes.Yellow)
                        {
                            run.Foreground = culoareTextVerset;
                        }
                    }
                }
            }

            // Reîncarcă versetele dacă avem o carte și capitol selectat
            if (versetCarteCurenta != null && versetCapitolCurent > 0)
            {
                IncarcaVersete(versetCarteCurenta, versetCapitolCurent);
            }
        }


        private void btnCautare_Click(object sender, RoutedEventArgs e)
        {
            string cuvant = txtCautare.Text.Trim();
            if (string.IsNullOrWhiteSpace(cuvant))
            {
                MessageBox.Show("Scrie un cuvânt pentru căutare.");
                return;
            }

            CautaVersete(cuvant);
        }




        private List<Inline> ParseTextCuTagRosu(string text)
        {
            var inlines = new List<Inline>();

            var pattern = @"<\s*ROSU\s*>([\s\S]*?)<\s*/\s*ROSU\s*>";

            var regex = new Regex(pattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);

            int lastIndex = 0;

            foreach (Match match in regex.Matches(text))
            {
                // textul dinainte
                if (match.Index > lastIndex)
                {
                    string normal = text.Substring(lastIndex, match.Index - lastIndex);
                    inlines.Add(new WpfRun(normal));
                }

                // textul dintre taguri
                string rosu = match.Groups[1].Value;
                inlines.Add(new WpfRun(rosu)
                {
                    Foreground = Brushes.Red,
                    FontWeight = FontWeights.Bold
                });

                lastIndex = match.Index + match.Length;
            }

            // textul de după ultimul tag
            if (lastIndex < text.Length)
            {
                string final = text.Substring(lastIndex);
                inlines.Add(new WpfRun(final));
            }

            return inlines;
        }








        private void IncarcaVerseteNormale()
        {
            listVersete.Items.Clear();

            var cmd = new SQLiteCommand(@"
        SELECT Carte, Capitol, Verset, Text 
        FROM Versete 
        ORDER BY CarteOrdine, Capitol, Verset", conn);

            var reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                string carte = reader.GetString(0);
                int capitol = reader.GetInt32(1);
                int verset = reader.GetInt32(2);
                string text = reader.GetString(3);

                var tb = new TextBlock
                {
                    TextWrapping = TextWrapping.Wrap,
                    FontSize = marimeTextCurenta,
                    FontFamily = new WpfFontFamily("Segoe UI"),
                    Foreground = culoareTextVerset,
                    Text = $"{carte} {capitol}:{verset} — {text}"
                };

                listVersete.Items.Add(new ListBoxItem
                {
                    Content = tb,
                    Foreground = culoareTextVerset
                });
            }
        }



        private List<Inline> DescompuneTextCuEvidentiere(string text, string cuvant, Brush culoare)
        {
            var inlines = new List<Inline>();

            if (string.IsNullOrWhiteSpace(cuvant))
            {
                // Tot textul e normal, setăm explicit Foreground
                inlines.Add(new WpfRun(text) { Foreground = culoare });
                return inlines;
            }

            var regex = new Regex(Regex.Escape(cuvant), RegexOptions.IgnoreCase);
            int index = 0;

            foreach (Match match in regex.Matches(text))
            {
                // Textul normal înainte de cuvântul evidențiat
                if (match.Index > index)
                {
                    string normal = text.Substring(index, match.Index - index);
                    inlines.Add(new WpfRun(normal)
                    {
                        Foreground = culoare
                    });
                }

                // Cuvântul evidențiat
                string evid = text.Substring(match.Index, match.Length);
                inlines.Add(new WpfRun(evid)
                {
                    Foreground = Brushes.Black, // text negru pe galben
                    Background = Brushes.Yellow,
                    FontWeight = FontWeights.Bold
                });

                index = match.Index + match.Length;
            }

            // Textul rămas după ultimul cuvânt evidențiat
            if (index < text.Length)
            {
                string rest = text.Substring(index);
                inlines.Add(new WpfRun(rest)
                {
                    Foreground = culoare
                });

            }

            return inlines;
        }


        private void btnCapitolAnterior_Click(object sender, RoutedEventArgs e)
        {
            if (sintetizatorGlobal != null && sintetizatorGlobal.State == SynthesizerState.Speaking)
            {
                sintetizatorGlobal.SpeakAsyncCancelAll();
            }

            if (versetCapitolCurent > 1)
            {
                versetCapitolCurent--;
                lblPozitieCurenta.Content = $"📖 {versetCarteCurenta} {versetCapitolCurent}";
                ActualizeazaTitlu();
                IncarcaVersete(versetCarteCurenta, versetCapitolCurent);

                // 🔹 doar aici
            }
            else
            {
                btnCapitolAnterior.ToolTip = new ToolTip
                {
                    Background = Brushes.LightYellow,
                    Padding = new Thickness(6),
                    Content = new TextBlock
                    {
                        Text = "Ești deja la primul capitol 📘",
                        Foreground = Brushes.DarkRed,
                        FontSize = 14,
                        FontWeight = FontWeights.Bold
                    }
                };

                ToolTipService.SetInitialShowDelay(btnCapitolAnterior, 0);
                ToolTipService.SetShowDuration(btnCapitolAnterior, 4000);
            }

        }

        private void btnCapitolUrmator_Click(object sender, RoutedEventArgs e)
        {
            if (sintetizatorGlobal != null && sintetizatorGlobal.State == SynthesizerState.Speaking)
            {
                sintetizatorGlobal.SpeakAsyncCancelAll();
            }

            var cmd = new SQLiteCommand("SELECT MAX(Capitol) FROM Versete WHERE Carte = @carte", conn);
            cmd.Parameters.AddWithValue("@carte", versetCarteCurenta);
            var maxCapitol = Convert.ToInt32(cmd.ExecuteScalar());

            if (versetCapitolCurent < maxCapitol)
            {
                versetCapitolCurent++;
                lblPozitieCurenta.Content = $"📖 {versetCarteCurenta} {versetCapitolCurent}";
                ActualizeazaTitlu();
                IncarcaVersete(versetCarteCurenta, versetCapitolCurent);
                // 🔹 doar aici
            }
            else
            {
                btnCapitolUrmator.ToolTip = new ToolTip
                {
                    Background = Brushes.LightYellow,
                    Padding = new Thickness(6),
                    Content = new TextBlock
                    {
                        Text = "Ești deja la ultimul capitol 📘",
                        Foreground = Brushes.DarkRed,
                        FontSize = 14,
                        FontWeight = FontWeights.Bold
                    }
                };

                ToolTipService.SetInitialShowDelay(btnCapitolUrmator, 0);
                ToolTipService.SetShowDuration(btnCapitolUrmator, 4000);
            }

        }


        //schita de aici in jos
        private void btnModSchita_Click(object sender, RoutedEventArgs e)
        {
          

            // Toggle Schiță
            modSchitaActiv = !modSchitaActiv;
            panouSchita.Visibility = modSchitaActiv ? Visibility.Visible : Visibility.Collapsed;

            if (ColSchita != null && ColCitire != null && ColMeniu != null)
            {
                if (modSchitaActiv)
                {
                    ColSchita.Width = new GridLength(2, GridUnitType.Star);
                    ColCitire.Width = new GridLength(3, GridUnitType.Star);
                    ColMeniu.Width = new GridLength(1, GridUnitType.Star);
                }
                else
                {
                    ColSchita.Width = new GridLength(0);
                    ColCitire.Width = new GridLength(4, GridUnitType.Star);
                    ColMeniu.Width = new GridLength(1, GridUnitType.Star);
                }
                gridPrincipal.UpdateLayout();
            }

            if (popup != null)
                popup.IsOpen = false;

            btnModSchita.Content = modSchitaActiv ? "Ascunde Schița ✏️" : "Arată Schița ✏️";

            btnModSchita.ToolTip = new ToolTip
            {
                Background = Brushes.LightYellow,
                Padding = new Thickness(6),
                Content = new TextBlock
                {
                    Text = modSchitaActiv ? "Schița este activă" : "Schița este ascunsă",
                    Foreground = modSchitaActiv ? Brushes.Green : Brushes.DarkBlue,
                    FontSize = 14,
                    FontWeight = FontWeights.Bold
                }
            };

            ToolTipService.SetInitialShowDelay(btnModSchita, 0);
            ToolTipService.SetShowDuration(btnModSchita, 4000);
            ToolTipService.SetBetweenShowDelay(btnModSchita, 0);

            // 🔹 Reîncarcă lista de versete ca să apară/dispară butonul ➕
            if (!string.IsNullOrEmpty(cuvantCurentCautat))
            {
                regex = new Regex(Regex.Escape(cuvantCurentCautat), RegexOptions.IgnoreCase);
                CautaVersete(cuvantCurentCautat);

                MessageBox.Show(
                    $"Lista de versete a fost reîncărcată pentru cuvântul: {cuvantCurentCautat}\n" +
                    $"Versete găsite: {totalVerseteGasite}, apariții: {totalAparitii}",
                    "BibliaApp - Mod Schiță",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information
                );
            }
            else if (!string.IsNullOrEmpty(versetCarteCurenta) && versetCapitolCurent > 0)
            {
                IncarcaVersete(versetCarteCurenta, versetCapitolCurent);
            }
        }


        private void CautaVersete(string cuvant)
        {
           

            listVersete.Items.Clear();
            totalAparitii = 0;
            totalVerseteGasite = 0;

            cuvant = cuvant.Trim().ToLower();
            cuvant = cuvant
    .Replace("â", "a")
    .Replace("ă", "a")
    .Replace("î", "i")
    .Replace("ș", "s")
    .Replace("ş", "s")
    .Replace("ț", "t")
    .Replace("ţ", "t");

            // 🔥 Labelul SUS (activăm și afișăm cuvântul căutat)
            lblRezultateCautare.Visibility = Visibility.Visible;
            lblRezultateCautare.Content = $"🔍 Căutare: {cuvant}";

            using (var cmd = new SQLiteCommand(
                "SELECT Carte, Capitol, Verset, TextNormalizat FROM Versete", conn))
            {
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string carte = reader.GetString(0);
                        int capitol = reader.GetInt32(1);
                        int verset = reader.GetInt32(2);
                        string text = reader.GetString(3);

                        string textNorm = text.ToLower();
                        string[] cuvinte = textNorm.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        int aparitii = cuvinte.Count(w => w == cuvant);

                        if (aparitii == 0)
                            continue;

                        totalVerseteGasite++;
                        totalAparitii += aparitii;

                        var tb = new TextBlock
                        {
                            TextWrapping = TextWrapping.Wrap,
                            FontSize = marimeTextCurenta,
                            FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                            Foreground = culoareTextVerset
                        };

                        tb.Inlines.Add(new Run($"{carte} {capitol}:{verset} — ")
                        {
                            FontWeight = FontWeights.Bold,
                            Foreground = culoareTextVerset
                        });

                        foreach (var inline in EvidentiazaCuvant(text, cuvant))
                            tb.Inlines.Add(inline);

                        listVersete.Items.Add(new ListBoxItem
                        {
                            Content = tb,
                            HorizontalContentAlignment = HorizontalAlignment.Stretch,
                            Foreground = culoareTextVerset
                        });
                    }
                }
            }

            if (totalVerseteGasite == 0)
            {
                lblRezultateCautare.Foreground = Brushes.Red;
                lblRezultateCautare.Content = $"⚠️ Nicio apariție pentru „{cuvant}”";
            }
            else
            {
                lblRezultateCautare.Foreground = Brushes.DarkGreen;
                lblRezultateCautare.Content =
                    $"📖 {totalVerseteGasite} versete • 🔎 {totalAparitii} cuvinte pentru „{cuvant}”";
            }


        }














        private void btnCauta_Click(object sender, RoutedEventArgs e)
        {
            

            cuvantCurentCautat = txtCautare.Text;
            regex = new Regex(Regex.Escape(cuvantCurentCautat), RegexOptions.IgnoreCase);
            CautaVersete(cuvantCurentCautat);
        }


        private void BtnAdaugaLaSchita_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is VersetSchita v)
            {
               
                AdaugaLaSchita(v);

                btn.Content = "✅";
                btn.IsEnabled = true;

                // confirmare imediată cu Popup (toast)
                var popup = new Popup
                {
                    
                    PlacementTarget = btn,
                    Placement = PlacementMode.Right,
                    StaysOpen = false,
                    Child = new WpfBorder
                    {
                        Background = Brushes.LightYellow,
                        Padding = new Thickness(6),
                        Child = new TextBlock
                        {
                            Text = "Versetul a fost adăugat în schiță",
                            Foreground = Brushes.Green,
                            FontSize = 16,
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


                    // acum delay‑urile funcționează corect
                    ToolTipService.SetInitialShowDelay(btn, 0);   // apare imediat
                    ToolTipService.SetShowDuration(btn, 5000);   // rămâne 5 secunde
                    ToolTipService.SetBetweenShowDelay(btn, 100);
                };
                timer.Start();
            }
        }









        private void AfiseazaSchita()
        {
            panouSchita.Visibility = Visibility.Visible;
            listaSchita.Children.Clear();

            for (int i = 0; i < schitaCurenta.Count; i++)
            {
                var v = schitaCurenta[i];

                // Textul versetului
                // Textul versetului (cu roșu pentru Isus)
                var tbVerset = new TextBlock
                {
                    TextWrapping = TextWrapping.Wrap,
                    FontWeight = FontWeights.SemiBold,
                    FontSize = marimeFontSchita + 2,
                    FontFamily = new WpfFontFamily(fontSchita),
                    Margin = new Thickness(0, 6, 0, 2)
                };

                // 🔹 Adaugă referința (Carte, Capitol, Verset)
                tbVerset.Inlines.Add(new WpfRun($"{v.Carte} {v.Capitol}:{v.Verset} — ") { Foreground = culoareTextVerset });

                // 🔹 Parsează textul cu roșu
                foreach (var inline in ParseCuvinteleLuiIsus(v.Text))
                {
                    tbVerset.Inlines.Add(inline);
                }


                // Câmp pentru comentariu
                var txtComentariu = new TextBox
                {
                    Text = v.Comentariu,
                    AcceptsReturn = true,
                    Height = 60,
                    TextWrapping = TextWrapping.Wrap,
                    Tag = v,
                    FontSize = marimeFontSchita, // 🔹 mai mic decât versetul
                    FontFamily = new WpfFontFamily("Segoe UI"),
                    Margin = new Thickness(0, 0, 0, 3)
                };

                txtComentariu.TextChanged += (s, e) =>
                {
                    if (s is TextBox t && t.Tag is VersetSchita vs)
                        vs.Comentariu = t.Text;
                };

                // Buton ștergere
                var btnSterge = new Button
                {
                    Content = "🗑️",
                    Width = 30,
                    Height = 30,
                    Margin = new Thickness(0, 0, 6, 0),
                    Tag = i
                };
                btnSterge.Click += (s, e) =>
                {
                    int index = (int)((Button)s).Tag;
                    schitaCurenta.RemoveAt(index);
                    AfiseazaSchita();
                };

                // Buton sus
                var btnSus = new Button
                {
                    Content = "⬆️",
                    Width = 30,
                    Height = 30,
                    Margin = new Thickness(0, 0, 6, 0),
                    Tag = i
                };
                btnSus.Click += (s, e) =>
                {
                    int index = (int)((Button)s).Tag;
                    if (index > 0)
                    {
                        var temp = schitaCurenta[index];
                        schitaCurenta[index] = schitaCurenta[index - 1];
                        schitaCurenta[index - 1] = temp;
                        AfiseazaSchita();
                    }
                };

                // Buton jos
                var btnJos = new Button
                {
                    Content = "⬇️",
                    Width = 30,
                    Height = 30,
                    Tag = i
                };
                btnJos.Click += (s, e) =>
                {
                    int index = (int)((Button)s).Tag;
                    if (index < schitaCurenta.Count - 1)
                    {
                        var temp = schitaCurenta[index];
                        schitaCurenta[index] = schitaCurenta[index + 1];
                        schitaCurenta[index + 1] = temp;
                        AfiseazaSchita();
                    }
                };

                // Grupare butoane
                var butoane = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Margin = new Thickness(0, 0, 0, 10)
                };
                butoane.Children.Add(btnSterge);
                butoane.Children.Add(btnSus);
                butoane.Children.Add(btnJos);

                // Adăugare în panou
                listaSchita.Children.Add(tbVerset);
                listaSchita.Children.Add(txtComentariu);
                listaSchita.Children.Add(butoane);
            }
        }

        private List<Inline> ParseCuvinteleLuiIsus(string text)
        {
            var inlines = new List<Inline>();

            // 1. Convertim ghilimelele în taguri interne
            text = text.Replace("„", "<QUOTE>").Replace("”", "</QUOTE>");

            // 2. Spargem textul după tagurile <ROSU>, </ROSU>, <QUOTE>, </QUOTE>
            var regex = new Regex("(<ROSU>|</ROSU>|<QUOTE>|</QUOTE>)");
            var parts = regex.Split(text);

            bool inRosu = false;
            bool inQuote = false;

            foreach (var part in parts)
            {
                if (part == "<ROSU>")
                {
                    inRosu = true;
                    continue;
                }
                if (part == "</ROSU>")
                {
                    inRosu = false;
                    continue;
                }
                if (part == "<QUOTE>")
                {
                    inQuote = true;
                    continue;
                }
                if (part == "</QUOTE>")
                {
                    inQuote = false;
                    continue;
                }

                // Text normal sau colorat
                var run = new WpfRun(part);

                if (inRosu || inQuote)
                {
                    run.Foreground = Brushes.Red;
                    run.FontWeight = FontWeights.Bold;
                }

                inlines.Add(run);
            }

            return inlines;
        }



        private void comboFontSchita_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (comboFontSchita.SelectedItem is ComboBoxItem item)
            {
                fontSchita = item.Content.ToString();
                AfiseazaSchita();
            }
        }

        private void comboFontSizeSchita_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (comboFontSizeSchita.SelectedItem is ComboBoxItem item && double.TryParse(item.Content.ToString(), out double size))
            {
                marimeFontSchita = size;
                AfiseazaSchita();
            }
        }

        // DOCX folosește half-points: 12pt => "24"
        private static string Pt(double pt) => ((int)(pt * 2)).ToString();

        // RunProperties cu font, mărime, bold, culoare opțională
        private OxmlRunProperties MakeRunProps(string fontName, double pt, bool bold = false, string hexColor = null)
        {
            // DOCX folosește half-points: 12pt => "24"
            string sizeVal = ((int)(pt * 2)).ToString();

            var rp = new OxmlRunProperties(
                new OxmlRunFonts { Ascii = fontName, HighAnsi = fontName },
                new OxmlFontSize { Val = sizeVal }
            );

            if (bold)
                rp.Append(new OxmlBold());

            if (!string.IsNullOrEmpty(hexColor))
                rp.Append(new OxmlColor { Val = hexColor });

            return rp;
        }


        // ParagraphProperties cu aliniere, spațiere Before/After și indentare opțională
        private OxmlParagraphProperties MakeParaProps(
    OxmlJustificationValues align,
    int beforeTwips = 0,
    int afterTwips = 0,
    int leftIndentTwips = 0)
        {
            var pp = new OxmlParagraphProperties(new OxmlJustification { Val = align });

            pp.Append(new OxmlSpacingBetweenLines
            {
                Before = beforeTwips.ToString(),
                After = afterTwips.ToString(),
                LineRule = DocumentFormat.OpenXml.Wordprocessing.LineSpacingRuleValues.Auto
            });

            if (leftIndentTwips > 0)
                pp.Append(new OxmlIndentation { Left = leftIndentTwips.ToString() });

            return pp;
        }


        // Împarte textul în bucăți normale/roșii bazat pe <ROSU> ... </ROSU>
        private IEnumerable<(string Text, bool IsRosu)> SplitRosu(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) yield break;

            text = Regex.Replace(text, @"<s\d{3}>", "<ROSU>");
            text = Regex.Replace(text, @"</s\d{3}>", "</ROSU>");

            int idx = 0;
            while (idx < text.Length)
            {
                int start = text.IndexOf("<ROSU>", idx);
                if (start == -1)
                {
                    var rest = text.Substring(idx);
                    if (rest.Length > 0) yield return (rest, false);
                    break;
                }
                if (start > idx)
                {
                    var normal = text.Substring(idx, start - idx);
                    if (normal.Length > 0) yield return (normal, false);
                }
                int end = text.IndexOf("</ROSU>", start);
                if (end == -1)
                {
                    var rosuFinal = text.Substring(start + 6);
                    if (rosuFinal.Length > 0) yield return (rosuFinal, true);
                    break;
                }
                var rosu = text.Substring(start + 6, end - (start + 6));
                if (rosu.Length > 0) yield return (rosu, true);
                idx = end + 7;
            }
        }


        // Paragraf de titlu (centru, bold, spațiere fixă)
        private OxmlParagraph MakeTitlu(string titlu, string fontName, double ptTitlu)
        {
            var p = new OxmlParagraph(
                MakeParaProps(OxmlJustificationValues.Center, beforeTwips: 240, afterTwips: 120)
            );

            var r = new OxmlRun(
                MakeRunProps(fontName, ptTitlu, bold: true),
                new OxmlText(titlu) { Space = SpaceProcessingModeValues.Preserve }
            );

            p.Append(r);
            return p;
        }


        // Paragraf de verset (stânga, referință bold, text cu <ROSU>)
        private OxmlParagraph MakeVerset(VersetSchita v, string fontName, double ptVerset)
        {
            var p = new OxmlParagraph(
                MakeParaProps(OxmlJustificationValues.Left, beforeTwips: 120, afterTwips: 60)
            );

            // Referința bold (carte cap:verset — )
            p.Append(new OxmlRun(
                MakeRunProps(fontName, ptVerset, bold: true),
                new OxmlText($"{v.Carte} {v.Capitol}:{v.Verset} — ") { Space = SpaceProcessingModeValues.Preserve }
            ));

            // Textul versetului, bucăți normale vs roșu
            foreach (var part in SplitRosu(v.Text))
            {
                var rp = MakeRunProps(fontName, ptVerset, bold: true, hexColor: part.IsRosu ? "FF0000" : null);
                p.Append(new OxmlRun(rp, new OxmlText(part.Text) { Space = SpaceProcessingModeValues.Preserve }));
            }

            return p;
        }


        // Comentarii în paragrafe separate, numerotate și indentate
        private IEnumerable<OxmlParagraph> MakeComentariuParagrafe(string comentariu, string fontName, double ptComentariu)
        {
            var linii = (comentariu ?? string.Empty).Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            int nr = 1;

            foreach (var linie in linii)
            {
                var p = new OxmlParagraph(
                    MakeParaProps(OxmlJustificationValues.Left, beforeTwips: 160, afterTwips: 40, leftIndentTwips: 360)
                );

                var run = new OxmlRun(
                    MakeRunProps(fontName, ptComentariu),
                    new OxmlText($"{nr}. {linie.Trim()}") { Space = SpaceProcessingModeValues.Preserve }
                );

                p.Append(run);
                yield return p;
                nr++;
            }
        }

        private List<string> ImpartireComentariuPeLinii(string text)
        {
            var linii = (text ?? string.Empty).Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            var rezultat = new List<string>();
            int nr = 1;

            foreach (var linie in linii)
            {
                rezultat.Add($"{nr}. {linie.Trim()}");
                nr++;
            }

            return rezultat;
        }


        private List<string> ImpartireComentariuInLiniiNumerotate(string text, XFont font, XGraphics gfx, double maxWidth)
        {
            var rezultat = new List<string>();
            if (string.IsNullOrWhiteSpace(text)) return rezultat;

            var cuvinte = text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var linieCurenta = new StringBuilder();
            int nr = 1;

            foreach (var cuvant in cuvinte)
            {
                string test = linieCurenta.Length == 0 ? cuvant : linieCurenta + " " + cuvant;
                double latime = gfx.MeasureString($"{nr}. {test}", font).Width;

                if (latime > maxWidth || test.Length > 80)
                {
                    if (linieCurenta.Length > 0)
                    {
                        rezultat.Add($"{nr}. {linieCurenta}");
                        nr++;
                        linieCurenta.Clear();
                    }

                    // dacă cuvântul e prea lung, îl rupem forțat
                    double latimeCuvant = gfx.MeasureString($"{nr}. {cuvant}", font).Width;
                    if (latimeCuvant > maxWidth)
                    {
                        int start = 0;
                        while (start < cuvant.Length)
                        {
                            int len = 1;
                            while (start + len <= cuvant.Length)
                            {
                                string chunk = cuvant.Substring(start, len);
                                double w = gfx.MeasureString($"{nr}. {chunk}", font).Width;
                                if (w > maxWidth) { len--; break; }
                                len++;
                            }
                            if (len <= 0) len = 1;

                            // verificare de siguranță
                            int lungimeMaxima = cuvant.Length - start;
                            if (len > lungimeMaxima) len = lungimeMaxima;

                            rezultat.Add($"{nr}. {cuvant.Substring(start, len)}");
                            nr++;
                            start += len;
                        }

                    }
                    else
                    {
                        linieCurenta.Append(cuvant);
                    }
                }
                else
                {
                    linieCurenta.Clear();
                    linieCurenta.Append(test);
                }
            }

            if (linieCurenta.Length > 0)
                rezultat.Add($"{nr}. {linieCurenta}");

            return rezultat;
        }



        // Dimensiuni pagină în twips (Word/OpenXML)
        private (UInt32Value W, UInt32Value H) GetPageSizeTwips(string formatHartie)
        {
            switch (formatHartie)
            {
                case "A4": return ((UInt32Value)11906U, (UInt32Value)16838U);
                case "A5": return ((UInt32Value)8391U, (UInt32Value)11906U);
                case "Letter": return ((UInt32Value)12240U, (UInt32Value)15840U);
                case "Legal": return ((UInt32Value)12240U, (UInt32Value)20160U);
                case "A6": return ((UInt32Value)5940U, (UInt32Value)8100U);
                default: return ((UInt32Value)11906U, (UInt32Value)16838U); // A4 fallback
            }
        }


        private void btnSalveazaPredica_Click(object sender, RoutedEventArgs e)
        {
            string format = ((ComboBoxItem)formatComboBox.SelectedItem)?.Content?.ToString();
            string formatHartie = ((ComboBoxItem)comboFormatHartie.SelectedItem)?.Content?.ToString() ?? "A4";
            string titlu = txtTitluPredica.Text.Trim();

            if (string.IsNullOrWhiteSpace(titlu))
            {
                MessageBox.Show("Scrie un titlu pentru predică.");
                return;
            }

            if (schitaCurenta == null || schitaCurenta.Count == 0)
            {
                MessageBox.Show("Predica nu are conținut.");
                return;
            }

            if (string.IsNullOrWhiteSpace(versetCarteCurenta) || versetCapitolCurent <= 0)
            {
                MessageBox.Show("Contextul (carte/capitol) nu este setat. Deschide un capitol înainte de salvare.");
                return;
            }

            string titluCurat = string.Concat(titlu.Split(Path.GetInvalidFileNameChars()));
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // Conținut pentru export TXT/PDF/DOCX
            string continut = string.Join("\n\n", schitaCurenta.Select(v =>
                $"{v.Carte} {v.Capitol}:{v.Verset} — {v.Text}\n{v.Comentariu}"));

            var schita = new SchitaPredica
            {
                Carte = versetCarteCurenta,
                Capitol = versetCapitolCurent,
                Titlu = titlu,
                Versete = schitaCurenta.ToList()
            };

            try
            {
                // 1) JSON pentru re-editare
                string fileJson = Path.Combine(desktop, $"{titluCurat}.json");
                var json = System.Text.Json.JsonSerializer.Serialize(schita);
                File.WriteAllText(fileJson, json);

                // 2) Export final TXT/PDF/DOCX
                string fileNameFinal = Path.Combine(desktop, $"{titluCurat}.{format?.ToLower()}");

                if (string.Equals(format, "TXT", StringComparison.OrdinalIgnoreCase))
                {
                    File.WriteAllText(fileNameFinal, $"{titlu}\n\n{continut}");
                    MessageBox.Show($"Salvat pe Desktop:\n- {Path.GetFileName(fileJson)} (JSON)\n- {Path.GetFileName(fileNameFinal)} (TXT)");
                }
                else if (string.Equals(format, "PDF", StringComparison.OrdinalIgnoreCase))
                {
                    var doc = new PdfDocument();
                    var page = doc.AddPage();

                    // Alegerea formatului de hârtie
                    switch (formatHartie)
                    {
                        case "A6":
                            page.Width = 297;   // 105 mm ≈ 297 puncte
                            page.Height = 420;  // 148 mm ≈ 420 puncte
                            break;
                        case "A5": page.Size = PdfSharp.PageSize.A5; break;
                        case "Letter": page.Size = PdfSharp.PageSize.Letter; break;
                        case "Legal": page.Size = PdfSharp.PageSize.Legal; break;
                        default: page.Size = PdfSharp.PageSize.A4; break;
                    }

                    var gfx = XGraphics.FromPdfPage(page);
                    var tf = new XTextFormatter(gfx);

                    // Fonturi adaptate pentru A6 (mai mici ca să încapă)
                    var fontTitlu = new XFont(fontSchita, marimeFontSchita + 1, XFontStyleEx.Bold);
                    var fontVerset = new XFont(fontSchita, marimeFontSchita - 2, XFontStyleEx.Bold);
                    var fontComentariu = new XFont(fontSchita, marimeFontSchita - 3, XFontStyleEx.Regular);

                    // Margini mai mici pentru A6
                    double x = 20;
                    double y = 20; // începe mai sus
                    double width = page.Width - 40;
                    double height = page.Height - 40;

                    // Titlu
                    tf.DrawString(titlu, fontTitlu, XBrushes.DarkBlue,
                        new XRect(x, y, width, 25), XStringFormats.TopLeft);
                    y += fontTitlu.GetHeight() + 2;

                    foreach (var v in schitaCurenta)
                    {
                        var versetText = $"{v.Carte} {v.Capitol}:{v.Verset} — {v.Text}";
                        var liniiVerset = ImpartireVersetPeLinii(versetText, fontVerset, gfx, width);

                        double inaltimeVerset = liniiVerset.Count * (fontVerset.GetHeight() + 5);
                        y += fontVerset.GetHeight() + 1;

                        if (y + inaltimeVerset > page.Height - 25)
                        {
                            page = doc.AddPage();
                            gfx = XGraphics.FromPdfPage(page);
                            y = 20; // începe sus pe noua pagină
                        }

                        foreach (var linie in liniiVerset)
                        {
                            gfx.DrawString(linie, fontVerset, XBrushes.Black,
                                new XRect(x, y, width, 25), XStringFormats.TopLeft);
                            y += fontVerset.GetHeight();
                        }

                        y += 2; // spațiu între verset și comentariu

                        var liniiComentariu = ImpartireComentariuPeLinii(v.Comentariu, fontComentariu, gfx, width - 20);

                        foreach (var linie in liniiComentariu)
                        {
                            gfx.DrawString(linie, fontComentariu, XBrushes.Black,
                                new XRect(x + 10, y, width - 20, 25), XStringFormats.TopLeft);
                            y += fontComentariu.GetHeight();

                            if (y > page.Height - 25)
                            {
                                page = doc.AddPage();
                                gfx = XGraphics.FromPdfPage(page);
                                y = 20;
                            }
                        }

                        y += 2; // spațiu între versete
                    }

                    doc.Save(fileNameFinal);
                    MessageBox.Show($"Salvat pe Desktop:\n- {Path.GetFileName(fileNameFinal)} (PDF A6)");
                }


                else if (string.Equals(format, "DOCX", StringComparison.OrdinalIgnoreCase))
                {
                    fileNameFinal = Path.Combine(desktop, $"{titluCurat}.docx");

                    // Format hârtie ales din combo (ai deja variabila formatHartie)
                    string formatHartieSelectat = formatHartie;

                    // Dimensiuni A6 în twips (1 twip = 1/1440 inch)
                    // A6 = 105 x 148 mm ≈ 297 x 420 puncte PDF ≈ 4252 x 5952 twips
                    var (pw, ph) = (4252, 5952);

                    // Fonturi adaptate pentru A6 (mai mici ca să încapă)
                    string fontTitlu = fontSchita;
                    double ptTitlu = marimeFontSchita + 1;
                    string fontVerset = fontSchita;
                    double ptVerset = marimeFontSchita - 2;
                    string fontComentariu = fontSchita;
                    double ptComentariu = marimeFontSchita - 3;

                    using (var docx = WordprocessingDocument.Create(fileNameFinal, WordprocessingDocumentType.Document))
                    {
                        var mainPart = docx.AddMainDocumentPart();
                        mainPart.Document = new Document();
                        var body = new Body();

                        // 1) Titlu
                        body.Append(MakeTitlu(titlu, fontTitlu, ptTitlu));

                        // 2) Versete + comentarii
                        foreach (var v in schitaCurenta)
                        {
                            body.Append(MakeVerset(v, fontVerset, ptVerset));

                            if (!string.IsNullOrWhiteSpace(v.Comentariu))
                            {
                                foreach (var pComent in MakeComentariuParagrafe(v.Comentariu, fontComentariu, ptComentariu))
                                    body.Append(pComent);
                            }
                        }

                        // 3) SectionProperties LA FINALUL Body (obligatoriu în Word)
                        var sectPr = new SectionProperties(
                         new PageSize { Width = 4252, Height = 5952, Orient = PageOrientationValues.Portrait },
                          new PageMargin { Top = 720, Right = 720, Bottom = 720, Left = 720 } // 0.5 inch margini
                        );

                        body.Append(sectPr);

                        mainPart.Document.Append(body);
                        mainPart.Document.Save();
                    }

                    MessageBox.Show($"Salvat pe Desktop:\n- {Path.GetFileName(fileNameFinal)} (DOCX A6)");
                }




                else
                {
                    MessageBox.Show("Alege formatul: TXT, PDF sau DOCX.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Eroare la salvare: " + ex.Message);
            }
        }
        // pentru VERSET (fără numerotare)
        private List<string> ImpartireVersetPeLinii(string text, XFont font, XGraphics gfx, double maxWidth)
        {
            var rezultat = new List<string>();
            if (string.IsNullOrWhiteSpace(text)) return rezultat;

            var paragrafe = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var paragraf in paragrafe)
            {
                var tokens = paragraf.Split(new[] { ' ' }, StringSplitOptions.None).Where(t => t.Length > 0).ToList();
                string linieCurenta = "";

                foreach (var token in tokens)
                {
                    string test = string.IsNullOrEmpty(linieCurenta) ? token : linieCurenta + " " + token;
                    var size = gfx.MeasureString(test, font);

                    if (size.Width <= maxWidth)
                    {
                        linieCurenta = test;
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(linieCurenta))
                        {
                            rezultat.Add(linieCurenta);
                            linieCurenta = "";
                        }
                        foreach (var segment in RupeTokenPeLatime(token, font, gfx, maxWidth))
                            rezultat.Add(segment);
                    }
                }

                if (!string.IsNullOrEmpty(linieCurenta))
                    rezultat.Add(linieCurenta);
            }

            return rezultat;
        }

        // pentru COMENTARIU (cu numerotare)
        private List<string> ImpartireComentariuPeLinii(string text, XFont font, XGraphics gfx, double maxWidth)
        {
            var rezultat = new List<string>();
            if (string.IsNullOrWhiteSpace(text)) return rezultat;

            var paragrafe = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            int nr = 1;

            foreach (var paragraf in paragrafe)
            {
                var tokens = paragraf.Split(new[] { ' ' }, StringSplitOptions.None).Where(t => t.Length > 0).ToList();
                string linieCurenta = "";

                foreach (var token in tokens)
                {
                    string test = string.IsNullOrEmpty(linieCurenta) ? token : linieCurenta + " " + token;
                    var size = gfx.MeasureString(test, font);

                    if (size.Width <= maxWidth)
                    {
                        linieCurenta = test;
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(linieCurenta))
                        {
                            rezultat.Add($"{nr}. {linieCurenta}");
                            nr++;
                            linieCurenta = "";
                        }
                        foreach (var segment in RupeTokenPeLatime(token, font, gfx, maxWidth))
                        {
                            rezultat.Add($"{nr}. {segment}");
                            nr++;
                        }
                    }
                }

                if (!string.IsNullOrEmpty(linieCurenta))
                {
                    rezultat.Add($"{nr}. {linieCurenta}");
                    nr++;
                }
            }

            return rezultat;
        }
        private IEnumerable<string> RupeTokenPeLatime(string token, XFont font, XGraphics gfx, double maxWidth)
        {
            int start = 0;
            while (start < token.Length)
            {
                int low = 1, high = token.Length - start, best = 1;

                while (low <= high)
                {
                    int mid = (low + high) / 2;
                    string probe = token.Substring(start, mid);
                    var size = gfx.MeasureString(probe, font);

                    if (size.Width <= maxWidth)
                    {
                        best = mid;
                        low = mid + 1;
                    }
                    else
                    {
                        high = mid - 1;
                    }
                }

                yield return token.Substring(start, best);
                start += best;
            }
        }

        private string ExtrageTextDin(UIElement element)
        {
            var sb = new System.Text.StringBuilder();

            if (element is TextBlock tb)
                sb.AppendLine(tb.Text); // ✅ verset

            else if (element is Label lbl)
                sb.AppendLine(lbl.Content?.ToString()); // ✅ titlu

            else if (element is TextBox txt)
                sb.AppendLine(txt.Text); // ✅ comentariu

            else if (element is Panel panel) // StackPanel, Grid etc.
            {
                foreach (UIElement child in panel.Children)
                {
                    sb.Append(ExtrageTextDin(child));
                }
            }

            return sb.ToString();
        }



        private void btnTrimitePredica_Click(object sender, RoutedEventArgs e)
        {
            var sb = new System.Text.StringBuilder();

            foreach (UIElement element in listaSchita.Children)
            {
                ExtractTextRecursive(element, sb);
                sb.AppendLine(); // spațiu între blocuri
            }

            string continutPredica = sb.ToString().Trim();
            string titluPredica = txtTitluPredica.Text.Trim();

            var emailWindow = new EmailWindow(continutPredica)
            {
                Owner = this,
                Subiect = string.IsNullOrWhiteSpace(titluPredica) ? "Predică" : titluPredica
            };

            if (emailWindow.ShowDialog() == true)
            {
                string destinatar = emailWindow.Destinatar;
                string subiect = emailWindow.Subiect;
                string continut = emailWindow.Continut;

                string mailto = $"mailto:{destinatar}?subject={Uri.EscapeDataString(subiect)}&body={Uri.EscapeDataString(continut)}";
                var psi = new System.Diagnostics.ProcessStartInfo(mailto) { UseShellExecute = true };
                System.Diagnostics.Process.Start(psi);
            }
        }



        private void ExtractTextRecursive(UIElement element, StringBuilder sb)
        {
            if (element is Button)
                return; // ❌ ignorăm complet butoanele

            if (element is TextBlock tb)
            {
                if (!string.IsNullOrWhiteSpace(tb.Text))
                {
                    sb.AppendLine(tb.Text);
                }
                else
                {
                    foreach (Inline inline in tb.Inlines)
                    {
                        if (inline is System.Windows.Documents.Run run)
                            sb.Append(run.Text);
                        else
                            sb.Append(new TextRange(inline.ContentStart, inline.ContentEnd).Text);
                    }
                    sb.AppendLine();
                }
            }
            else if (element is TextBox txt)
            {
                sb.AppendLine(txt.Text);
            }
            else if (element is Label lbl)
            {
                sb.AppendLine(lbl.Content?.ToString());
            }
            else if (element is RichTextBox rtb)
            {
                var range = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
                sb.AppendLine(range.Text);
            }

            int childCount = VisualTreeHelper.GetChildrenCount(element);
            for (int i = 0; i < childCount; i++)
            {
                var child = VisualTreeHelper.GetChild(element, i) as UIElement;
                if (child != null)
                    ExtractTextRecursive(child, sb);
            }
        }















        private void TrimitePredicaPrinEmail(string destinatar, string formatPreferat)
        {
            string titlu = txtTitluPredica.Text.Trim();
            string titluCurat = string.Concat(titlu.Split(Path.GetInvalidFileNameChars()));
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // **Corpul textului (fallback)**
            string continut = string.Join("\n\n", schitaCurenta.Select(v =>
                $"{v.Carte} {v.Capitol}:{v.Verset} — {v.Text}\n{v.Comentariu}"));

            string atașamentCale = null;

            // **Dacă a cerut atașament, generăm fișierul**
            if (!string.IsNullOrWhiteSpace(formatPreferat))
            {
                string format = formatPreferat.Trim().ToUpperInvariant();
                try
                {
                    if (format == "TXT")
                    {
                        atașamentCale = Path.Combine(desktop, $"{titluCurat}.txt");
                        File.WriteAllText(atașamentCale, $"{titlu}\n\n{continut}");
                    }
                    else if (format == "PDF")
                    {
                        atașamentCale = Path.Combine(desktop, $"{titluCurat}.pdf");
                        ExportSchitaCaPDF(atașamentCale); // vezi funcția de mai jos
                    }
                    else if (format == "DOCX")
                    {
                        atașamentCale = Path.Combine(desktop, $"{titluCurat}.docx");
                        ExportSchitaCaDOCX(atașamentCale); // vezi funcția de mai jos
                    }
                    else
                    {
                        MessageBox.Show("Format necunoscut. Folosesc text în corpul emailului.");
                        atașamentCale = null;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Eroare la generarea atașamentului: " + ex.Message);
                    atașamentCale = null; // fallback la corp text
                }
            }

            // **Trimitere SMTP**
            TrimiteEmail(destinatar, $"Predica: {titlu}", $"{titlu}\n\n{continut}", atașamentCale);
        }
        private void ExportSchitaCaPDF(string cale)
        {
            var doc = new PdfDocument();
            var page = doc.AddPage();
            var gfx = XGraphics.FromPdfPage(page);
            var tf = new XTextFormatter(gfx);

            var fontTitlu = new XFont(fontSchita, marimeFontSchita + 1, XFontStyleEx.Bold);
            var fontVerset = new XFont(fontSchita, marimeFontSchita - 2, XFontStyleEx.Bold);
            var fontComentariu = new XFont(fontSchita, marimeFontSchita - 3, XFontStyleEx.Regular);

            double x = 40, y = 40, width = page.Width - 80;

            tf.DrawString(txtTitluPredica.Text.Trim(), fontTitlu, XBrushes.DarkBlue,
                new XRect(x, y, width, 25), XStringFormats.TopLeft);
            y += fontTitlu.GetHeight() + 6;

            foreach (var v in schitaCurenta)
            {
                var versetText = $"{v.Carte} {v.Capitol}:{v.Verset} — {v.Text}";
                var liniiVerset = ImpartireVersetPeLinii(versetText, fontVerset, gfx, width);
                foreach (var linie in liniiVerset)
                {
                    gfx.DrawString(linie, fontVerset, XBrushes.Black, new XRect(x, y, width, 25), XStringFormats.TopLeft);
                    y += fontVerset.GetHeight();
                }
                y += 2;
                var liniiComent = ImpartireComentariuPeLinii(v.Comentariu, fontComentariu, gfx, width - 20);
                foreach (var linie in liniiComent)
                {
                    gfx.DrawString(linie, fontComentariu, XBrushes.Black, new XRect(x + 10, y, width - 20, 25), XStringFormats.TopLeft);
                    y += fontComentariu.GetHeight();
                }
                y += 6;
                if (y > page.Height - 60)
                {
                    page = doc.AddPage();
                    gfx = XGraphics.FromPdfPage(page);
                    tf = new XTextFormatter(gfx);
                    y = 40;
                }
            }

            doc.Save(cale);
        }

        private void ExportSchitaCaDOCX(string cale)
        {
            using var docx = WordprocessingDocument.Create(cale, WordprocessingDocumentType.Document);
            var mainPart = docx.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();

            string fontTitlu = fontSchita;
            double ptTitlu = marimeFontSchita + 1;
            string fontVerset = fontSchita;
            double ptVerset = marimeFontSchita - 2;
            string fontComentariu = fontSchita;
            double ptComentariu = marimeFontSchita - 3;

            body.Append(MakeTitlu(txtTitluPredica.Text.Trim(), fontTitlu, ptTitlu));
            foreach (var v in schitaCurenta)
            {
                body.Append(MakeVerset(v, fontVerset, ptVerset));
                if (!string.IsNullOrWhiteSpace(v.Comentariu))
                {
                    foreach (var p in MakeComentariuParagrafe(v.Comentariu, fontComentariu, ptComentariu))
                        body.Append(p);
                }
            }

            var sectPr = new SectionProperties(
                new PageSize { Width = 11906U, Height = 16838U, Orient = PageOrientationValues.Portrait },
                new PageMargin { Top = 720, Right = 720, Bottom = 720, Left = 720 }
            );
            body.Append(sectPr);

            mainPart.Document.Append(body);
            mainPart.Document.Save();
        }
        private void TrimiteEmail(string destinatar, string subiect, string bodyText, string attachmentPath = null)
        {
            try
            {
                var mail = new MailMessage();
                mail.From = new MailAddress("adresa_ta@example.com", "BibliaApp");
                mail.To.Add(destinatar);
                mail.Subject = subiect;
                mail.Body = bodyText;

                if (!string.IsNullOrWhiteSpace(attachmentPath) && File.Exists(attachmentPath))
                {
                    mail.Attachments.Add(new Attachment(attachmentPath));
                }

                // SMTP: exemplu Gmail
                using var smtp = new SmtpClient("smtp.gmail.com", 587)
                {
                    Credentials = new NetworkCredential("adresa_ta@example.com", "PAROLA_SAU_PAROLA_APLICATIE"),
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network
                };

                // Feedback vizual înainte de trimitere
                btnTrimitePredica.Content = "Trimit... ⏳";
                btnTrimitePredica.IsEnabled = false;

                smtp.Send(mail);

                // Confirmare
                btnTrimitePredica.Content = "Trimis ✉️";
                btnTrimitePredica.IsEnabled = true;
                MessageBox.Show("Email trimis cu succes.");
            }
            catch (SmtpException ex)
            {
                btnTrimitePredica.Content = "Trimite prin email ✉️";
                btnTrimitePredica.IsEnabled = true;
                MessageBox.Show("Eroare SMTP: " + ex.Message);
            }
            catch (Exception ex)
            {
                btnTrimitePredica.Content = "Trimite prin email ✉️";
                btnTrimitePredica.IsEnabled = true;
                MessageBox.Show("Eroare la trimitere: " + ex.Message);
            }
        }
       


        private string GetFormatHartieSelectat()
        {
            if (comboFormatHartie.SelectedItem is ComboBoxItem item)
                return item.Content.ToString();
            return "A4";
        }



        private void btnIncarcaSchita_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Fișiere schiță (*.json)|*.json",
                Title = "Selectează schița salvată"
            };

            if (dialog.ShowDialog() != true) return;

            try
            {
                var json = File.ReadAllText(dialog.FileName);
                var schita = System.Text.Json.JsonSerializer.Deserialize<SchitaPredica>(json);

                if (schita == null)
                {
                    MessageBox.Show("Fișier JSON invalid: nu s-a putut deserializa.");
                    return;
                }

                if (string.IsNullOrWhiteSpace(schita.Carte) || schita.Capitol <= 0)
                {
                    MessageBox.Show("Fișierul nu conține context valid (carte/capitol).");
                    return;
                }

                // 1) Restaurăm contextul
                versetCarteCurenta = schita.Carte;
                versetCapitolCurent = schita.Capitol;
                txtTitluPredica.Text = schita.Titlu ?? "";

                // 2) Restaurăm lista versetelor din schiță
                schitaCurenta = schita.Versete ?? new List<VersetSchita>();

                // 3) Activăm modul Schiță și panoul
                // Activăm modul Schiță
                modSchitaActiv = true;
                panouSchita.Visibility = Visibility.Visible;

                if (ColSchita != null)
                {
                    ColSchita.Width = new GridLength(2, GridUnitType.Star);
                }


                // 4) Reîncărcăm capitolul potrivit
                IncarcaVersete(versetCarteCurenta, versetCapitolCurent);

                // 5) Reafișăm panoul schiței
                AfiseazaSchita();

                // 6) Actualizăm titlul lateral
                ActualizeazaTitlu();

                MessageBox.Show($"Încărcat: {versetCarteCurenta} {versetCapitolCurent} — {schitaCurenta.Count} elemente în schiță.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Eroare la încărcare: " + ex.Message);
            }
        }






        private void AfiseazaSchitaPentruEditare()
        {
            listVersete.Items.Clear();

            foreach (var v in schitaCurenta)
            {
                var tb = new TextBlock
                {
                    Text = $"{v.Carte} {v.Capitol}:{v.Verset} — {v.Text}",
                    TextWrapping = TextWrapping.Wrap,
                    FontSize = marimeTextCurenta,
                    Foreground = culoareTextVerset
                };

                var txtComentariu = new TextBox
                {
                    Text = v.Comentariu,
                    Width = 300,
                    Margin = new Thickness(6, 0, 0, 0),
                    Tag = v
                };
                txtComentariu.TextChanged += (s, e) =>
                {
                    if (s is TextBox box && box.Tag is VersetSchita vs)
                        vs.Comentariu = box.Text;
                };

                var stack = new StackPanel { Orientation = Orientation.Vertical };
                stack.Children.Add(tb);
                stack.Children.Add(txtComentariu);

                listVersete.Items.Add(new ListBoxItem { Content = stack });
            }
        }

        private void panouSchita_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            isDraggingSchita = true;
            startPointSchita = e.GetPosition(this);
            panouSchita.CaptureMouse();
        }

        private void panouSchita_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDraggingSchita)
            {
                Point currentPoint = e.GetPosition(this);
                double offsetX = currentPoint.X - startPointSchita.X;
                double offsetY = currentPoint.Y - startPointSchita.Y;

                Thickness currentMargin = panouSchita.Margin;
                panouSchita.Margin = new Thickness(
                    currentMargin.Left + offsetX,
                    currentMargin.Top + offsetY,
                    0, 0);

                startPointSchita = currentPoint;
            }
        }

        private void panouSchita_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            isDraggingSchita = false;
            panouSchita.ReleaseMouseCapture();
        }
        private void btnInchideSchita_Click(object sender, RoutedEventArgs e)
        {
            panouSchita.Visibility = Visibility.Collapsed;
            modSchitaActiv = false;

            if (ColSchita != null)
            {
                ColSchita.Width = new GridLength(0); // eliberează spațiul pentru zona de citire
            }
        }

        private void ExportaPredicaInFisier()
        {
            if (string.IsNullOrWhiteSpace(txtTitluPredica.Text))
            {
                MessageBox.Show("Scrie un titlu pentru predică.");
                return;
            }

            string titlu = txtTitluPredica.Text.Trim();
            string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Predici");
            Directory.CreateDirectory(folder);

            string caleFisier = Path.Combine(folder, $"{titlu}.txt");

            using (StreamWriter sw = new StreamWriter(caleFisier))
            {
                sw.WriteLine($"📝 Titlu: {titlu}");
                sw.WriteLine($"📅 Data: {DateTime.Now:yyyy-MM-dd HH:mm}");
                sw.WriteLine(new string('-', 40));

                foreach (var v in schitaCurenta)
                {
                    sw.WriteLine($"📖 {v.Carte} {v.Capitol}:{v.Verset}");
                    sw.WriteLine($"Text: {v.Text}");
                    sw.WriteLine($"💡 Idee: {v.Comentariu}");
                    sw.WriteLine();
                }
            }

            MessageBox.Show($"Predica a fost salvată pe Desktop în folderul:\n{folder}");
        }
        private void ExportaPredicaInPDF()
        {
            if (string.IsNullOrWhiteSpace(txtTitluPredica.Text))
            {
                MessageBox.Show("Scrie un titlu pentru predică.");
                return;
            }

            var dialog = new SaveFileDialog
            {
                Title = "Salvează predica ca PDF",
                Filter = "Fișier PDF (*.pdf)|*.pdf",
                FileName = txtTitluPredica.Text.Trim() + ".pdf"
            };

            if (dialog.ShowDialog() != true)
                return;

            var document = new PdfDocument();
            document.Info.Title = txtTitluPredica.Text.Trim();

            var page = document.AddPage();
            var gfx = XGraphics.FromPdfPage(page);

            var fontTitlu = new XFont(fontSchita, marimeFontSchita + 4, XFontStyleEx.Bold);
            var fontText = new XFont(fontSchita, marimeFontSchita, XFontStyleEx.Regular);






            double y = 40;
            gfx.DrawString($"📝 Titlu: {txtTitluPredica.Text.Trim()}", fontTitlu, XBrushes.Black,
                new XRect(40, y, page.Width - 80, page.Height), XStringFormats.TopLeft);
            y += 30;

            gfx.DrawString($"📅 Data: {DateTime.Now:yyyy-MM-dd HH:mm}", fontText, XBrushes.Black,
                new XRect(40, y, page.Width - 80, page.Height), XStringFormats.TopLeft);
            y += 30;

            foreach (var v in schitaCurenta)
            {
                string bloc = $"📖 {v.Carte} {v.Capitol}:{v.Verset}\nText: {v.Text}\n💡 Idee: {v.Comentariu}\n";
                var lines = bloc.Split('\n');

                foreach (var line in lines)
                {
                    gfx.DrawString(line, fontText, XBrushes.Black,
                        new XRect(40, y, page.Width - 80, page.Height), XStringFormats.TopLeft);
                    y += 20;

                    if (y > page.Height - 60)
                    {
                        page = document.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                        y = 40;
                    }
                }

                y += 10;
            }

            document.Save(dialog.FileName);
            MessageBox.Show("Predica a fost salvată cu succes în PDF!");
        }
        void SalveazaCaPDF()
        {
            var document = new PdfDocument();
            var page = document.AddPage();
            var gfx = XGraphics.FromPdfPage(page);

            var font = new XFont("LiberationSans", 14, XFontStyleEx.Regular);
            gfx.DrawString("Textul tău biblic sau tematic aici", font, XBrushes.Black, new XRect(40, 40, page.Width - 80, page.Height - 80), XStringFormats.TopLeft);

            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "salvare.pdf");
            document.Save(path);
            MessageBox.Show("Fișierul PDF a fost salvat pe Desktop.");
        }

        void SalveazaCaText()
        {
            var continut = "Textul tău biblic sau tematic aici";
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "salvare.txt");
            File.WriteAllText(path, continut);
            MessageBox.Show("Fișierul TXT a fost salvat pe Desktop.");
        }

        private void BtnAdaugaFavorite_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is VersetSchita verset)
            {
                var fereastra = new FereastraFavorite(favoritePeListe.Keys);
                if (fereastra.ShowDialog() == true)
                {
                    string listaAleasa = fereastra.ListaAleasa;
                    string comentariu = fereastra.ComentariuScris;

                    var versetCuComentariu = new VersetSchita
                    {
                        Carte = verset.Carte,
                        Capitol = verset.Capitol,
                        Verset = verset.Verset,
                        Text = verset.Text,
                        Comentariu = comentariu
                    };

                    AdaugaVersetLaLista(listaAleasa, versetCuComentariu);

                    favoritePeListe = IncarcaFavoriteDinBazaDeDate();

                    // 🔥 Rămânem în Scriptură după salvare
                    SeteazaMod(ModAfisare.Scriptura);

                    comboListeFavorite.SelectionChanged -= comboListeFavorite_SelectionChanged;
                    comboListeFavorite.ItemsSource = favoritePeListe.Keys.ToList();
                    comboListeFavorite.Text = listaAleasa;
                    comboListeFavorite.SelectionChanged += comboListeFavorite_SelectionChanged;
                }
            }
        }





        private void btnFavorite_Click(object sender, RoutedEventArgs e)
        {
            SeteazaMod(ModAfisare.Favorite);

            favoritePeListe = IncarcaFavoriteDinBazaDeDate();

            var f = new FereastraFavoriteView(favoritePeListe);
            f.Owner = this;
            f.ShowDialog();

            SeteazaMod(ModAfisare.Scriptura);

            favoritePeListe = IncarcaFavoriteDinBazaDeDate();
            comboListeFavorite.ItemsSource = favoritePeListe.Keys.ToList();

            // 🔥 Încarcă Scriptura doar dacă există carte și capitol valid
            if (!string.IsNullOrWhiteSpace(CarteCurenta) && CapitolCurent > 0)
            {
                IncarcaVersete(CarteCurenta, CapitolCurent);
            }
        }









        // dacă constructorul e FereastraFavorite(IEnumerable<string> numeListe)





        private void AfiseazaListaFavorite(string lista)
        {
            // 🔥 Intrăm aici DOAR dacă suntem în modul Favorite
            if (!EsteMod(ModAfisare.Favorite))
                return;

            // 🔥 Dacă lista e null sau goală, ieșim
            if (string.IsNullOrWhiteSpace(lista))
                return;

            // 🔥 Dacă favoritele nu sunt încărcate, ieșim
            if (favoritePeListe == null)
                return;

            // 🔥 Dacă lista nu există în dicționar, ieșim
            if (!favoritePeListe.ContainsKey(lista))
                return;

            // 🔥 Dacă lista e goală, ieșim
            if (favoritePeListe[lista].Count == 0)
                return;

            // 🔥 Curățăm UI-ul
            listVersete.Items.Clear();

            // 🔥 Afișăm favoritele în mod sigur
            foreach (var v in favoritePeListe[lista])
            {
                // afișare
            }

            // 🔥 Actualizăm eticheta
            lblPozitieCurenta.Content = $"⭐ Lista: {lista}";
        }


        private void BtnTrimiteLaSchitaDinFavorite_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is VersetSchita verset)
            {
                // Folosește metoda unică ce are deduplicare + reload + MessageBox
                AdaugaLaSchita(verset);

                // Dacă ești în modul schiță și capitolul vizualizat e același cu cel al versetului, reîncarcă vizual capitolul
                if (modSchitaActiv &&
                    versetCarteCurenta == verset.Carte &&
                    versetCapitolCurent == verset.Capitol)
                {
                    IncarcaVersete(versetCarteCurenta, versetCapitolCurent); // ✅ aici apare butonul ✅ în loc de ➕
                }
            }
        }

        private void comboListeFavorite_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 🔥 Intrăm aici DOAR dacă suntem în modul Favorite
            if (!EsteMod(ModAfisare.Favorite))
                return;

            // 🔥 Dacă nu e selectat nimic, ieșim
            if (comboListeFavorite.SelectedItem is not string lista)
                return;

            // 🔥 Dacă lista nu există în dicționar, ieșim
            if (!favoritePeListe.ContainsKey(lista))
                return;

            // 🔥 Afișăm favoritele în mod sigur
            AfiseazaListaFavorite(lista);
        }



        private void btnAdaugaInListaFavorite_Click(object sender, RoutedEventArgs e)
        {
            if (listVersete.SelectedItem is not ListBoxItem item || item.Content is not StackPanel stack)
                return;

            var btn = stack.Children.OfType<Button>().FirstOrDefault(b => b.Tag is VersetSchita);
            var verset = btn?.Tag as VersetSchita;
            if (verset == null)
                return;

            string lista = comboListeFavorite.Text.Trim();
            if (string.IsNullOrWhiteSpace(lista))
            {
                MessageBox.Show("Scrie un nume pentru listă.");
                return;
            }

            AdaugaVersetLaLista(lista, verset);

            favoritePeListe = IncarcaFavoriteDinBazaDeDate();

            // 🔥 RĂMÂNEM ÎN SCRIPTURĂ, NU INTRĂM ÎN FAVORITE
            SeteazaMod(ModAfisare.Scriptura);

            comboListeFavorite.SelectionChanged -= comboListeFavorite_SelectionChanged;
            comboListeFavorite.ItemsSource = favoritePeListe.Keys.ToList();
            comboListeFavorite.SelectedItem = lista;
            comboListeFavorite.SelectionChanged += comboListeFavorite_SelectionChanged;

            MessageBox.Show($"Verset adăugat în lista „{lista}”.");
        }





        private void btnStergeListaFavorite_Click(object sender, RoutedEventArgs e)
        {
            if (comboListeFavorite.SelectedItem is not string lista)
            {
                MessageBox.Show("Selectează o listă validă.");
                return;
            }

            var confirm = MessageBox.Show(
                $"Sigur vrei să ștergi lista „{lista}” cu toate versetele?",
                "Confirmare",
                MessageBoxButton.YesNo);

            if (confirm == MessageBoxResult.Yes)
            {
                // 🔥 Ștergem direct din baza de date
                StergeLista(lista);

                // 🔥 Reîncărcăm structura în memorie
                favoritePeListe = IncarcaFavoriteDinBazaDeDate();

                // 🔥 Actualizăm UI-ul
                comboListeFavorite.ItemsSource = favoritePeListe.Keys.ToList();
                comboListeFavorite.SelectedItem = null;

                listVersete.Items.Clear();
                lblPozitieCurenta.Content = "⭐ Selectează o listă de favorite";
            }
        }

        private void btnRedenumesteListaFavorite_Click(object sender, RoutedEventArgs e)
        {
            if (comboListeFavorite.SelectedItem is not string listaVeche)
            {
                MessageBox.Show("Selectează o listă validă.");
                return;
            }

            var nouaDenumire = Microsoft.VisualBasic.Interaction.InputBox(
                $"Scrie noul nume pentru lista „{listaVeche}”",
                "Redenumește lista",
                listaVeche);

            if (string.IsNullOrWhiteSpace(nouaDenumire))
                return;

            nouaDenumire = nouaDenumire.Trim();

            // 🔥 Verificăm dacă există deja o listă cu acest nume
            if (favoritePeListe.ContainsKey(nouaDenumire))
            {
                MessageBox.Show("Există deja o listă cu acest nume.");
                return;
            }

            // 🔥 Redenumim direct în baza de date
            RedenumesteLista(listaVeche, nouaDenumire);

            // 🔥 Reîncărcăm structura în memorie
            favoritePeListe = IncarcaFavoriteDinBazaDeDate();

            // 🔥 Actualizăm UI-ul
            comboListeFavorite.ItemsSource = favoritePeListe.Keys.ToList();
            comboListeFavorite.SelectedItem = nouaDenumire;

            AfiseazaListaFavorite(nouaDenumire);
        }

        private void listVersete_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            e.Handled = true;

            var eventArg = new MouseWheelEventArgs(e.MouseDevice, e.Timestamp, e.Delta)
            {
                RoutedEvent = UIElement.MouseWheelEvent,
                Source = sender
            };

            var parent = VisualTreeHelper.GetParent((DependencyObject)sender) as UIElement;
            parent?.RaiseEvent(eventArg);
        }

        private static T FindVisualChild<T>(DependencyObject obj) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(obj, i);
                if (child is T tChild)
                    return tChild;

                T childOfChild = FindVisualChild<T>(child);
                if (childOfChild != null)
                    return childOfChild;
            }
            return null;
        }
        public void CitesteText(string text)
        {
            sintetizatorGlobal?.SpeakAsyncCancelAll();

            sintetizatorGlobal = new SpeechSynthesizer();
            sintetizatorGlobal.SelectVoice("Microsoft Andrei"); // sau Irina dacă o ai

            // 🔧 Salvează textul complet pentru reluare
            textCititComplet = text;
            pozitieCurenta = 0;
            esteInPauza = false;

            sintetizatorGlobal.SpeakProgress += SintetizatorGlobal_SpeakProgress;
            sintetizatorGlobal.SpeakAsync(text);
        }
        private void btnCitesteCapitol_Click(object sender, RoutedEventArgs e)
        {
            string textCurat = ObtineTextCapitolCurat(versetCarteCurenta, versetCapitolCurent);

            if (!string.IsNullOrWhiteSpace(textCurat))
            {
                CitesteText(textCurat);
            }
            else
            {
                MessageBox.Show("Nu există versete în capitolul curent.",
                                "Listă goală",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
            }
        }




        private void btnOpresteCitirea_Click(object sender, RoutedEventArgs e)
        {
            sintetizatorGlobal?.SpeakAsyncCancelAll();
        }
        private void SintetizatorGlobal_SpeakProgress(object sender, SpeakProgressEventArgs e)
        {
            pozitieCurenta = e.CharacterPosition;
        }
        private void btnPauzaCitire_Click(object sender, RoutedEventArgs e)
        {
            if (sintetizatorGlobal != null)
            {
                sintetizatorGlobal.SpeakAsyncCancelAll();
                esteInPauza = true;
            }
        }
        private void btnReiaCitirea_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(textCititComplet))
            {
                MessageBox.Show("Nu există text salvat pentru reluare.", "Eroare", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (sintetizatorGlobal != null && esteInPauza && pozitieCurenta < textCititComplet.Length)
            {
                sintetizatorGlobal = new SpeechSynthesizer();
                sintetizatorGlobal.SelectVoice("Microsoft Andrei"); // sau Irina dacă o ai
                sintetizatorGlobal.SpeakProgress += SintetizatorGlobal_SpeakProgress;

                string textRamas = textCititComplet.Substring(pozitieCurenta);
                sintetizatorGlobal.SpeakAsync(textRamas);
                esteInPauza = false;
            }
        }
        private string ObtineTextCapitolCurat(string carte, int capitol)
        {
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "biblia.db");
            using var conn = new SQLiteConnection($"Data Source={path}");
            conn.Open();

            var cmd = new SQLiteCommand(
                "SELECT Text FROM Versete WHERE Carte = @carte AND Capitol = @capitol ORDER BY Verset", conn);

            cmd.Parameters.AddWithValue("@carte", carte);
            cmd.Parameters.AddWithValue("@capitol", capitol);

            using var reader = cmd.ExecuteReader();

            var sb = new StringBuilder();

            while (reader.Read())
            {
                string txt = reader.GetString(0);

                // elimină marcajele <ROSU>
                txt = txt.Replace("<ROSU>", "").Replace("</ROSU>", "");

                sb.Append(txt).Append(" ");
            }

            return sb.ToString().Trim();
        }

        //comentara 
        private void BtnComentariu_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is VersetSchita v)
            {
                var fereastra = new FereastraComentariu(conn, v.Carte, v.Capitol, v.Verset);
                fereastra.ShowDialog();
            }



            
        }


        private void AdaugaComentariuLaSchita(string blocText)
        {
            var bloc = new VersetSchita
            {
                Carte = versetCarteCurenta,
                Capitol = versetCapitolCurent,
                Verset = 0,
                Text = blocText
            };

            schitaCurenta.Add(bloc);

        }

        private void BtnStergeBlocSchita_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is int index)
            {
                schitaCurenta.RemoveAt(index);
                AfiseazaSchita();
            }

        }
        private void BtnCopiazaSchita_Click(object sender, RoutedEventArgs e)
        {
            var text = string.Join("\n\n", schitaCurenta.Select(b =>
     $"{b.Carte} {b.Capitol}:{b.Verset} — {b.Text}\n{b.Comentariu}"));

        }

        private bool AreComentariu(string carte, int capitol, int verset)
        {
            string carteNorm = NormalizareCarte(carte);

            var cmd = new SQLiteCommand(
                "SELECT COUNT(*) FROM Comentarii WHERE UPPER(Carte) = @carte AND Capitol = @capitol AND Verset = @verset",
                conn);

            cmd.Parameters.AddWithValue("@carte", carteNorm.ToUpper());
            cmd.Parameters.AddWithValue("@capitol", capitol);
            cmd.Parameters.AddWithValue("@verset", verset);

            var count = Convert.ToInt32(cmd.ExecuteScalar());
            return count > 0;
        }

        private bool ExistaFavorite(string carte, int capitol)
        {
            // verifică dacă există favorite pentru capitolul curent
            return favoritePeListe.Values.Any(lista =>
                lista.Any(v => v.Carte == carte && v.Capitol == capitol));
        }

        private bool ExistaComentarii(string carte, int capitol)
        {
            string carteNorm = NormalizareCarte(carte);

            var cmd = new SQLiteCommand(
                "SELECT COUNT(*) FROM Comentarii WHERE UPPER(Carte)=@carte AND Capitol=@capitol",
                conn);

            cmd.Parameters.AddWithValue("@carte", carteNorm.ToUpper());
            cmd.Parameters.AddWithValue("@capitol", capitol);

            return Convert.ToInt32(cmd.ExecuteScalar()) > 0;
        }



        private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (sender is ScrollViewer scrollViewer)
            {
                // mută scroll-ul în sus sau jos
                scrollViewer.ScrollToVerticalOffset(scrollViewer.VerticalOffset - e.Delta);
                e.Handled = true;
            }
        }



        // ✅ Actualizează titlul cu simbolurile ca indicatori de conținut
        private void ActualizeazaTitlu()
        {
            // ✅ Dacă nu ai selectat nicio carte
            if (string.IsNullOrEmpty(versetCarteCurenta) || versetCapitolCurent <= 0)
            {
                runTitlu.Text = "📖 Cărțile BIBLIEI";
                runStea.Text = "";

                var tbStea = (TextBlock)runStea.Parent;
                var sbStop = (Storyboard)tbStea.Resources["PulseStar"];
                sbStop.Stop();

                txtStea.ToolTip = null;
                return;
            }

            // ✅ Ai selectat o carte și un capitol
            runTitlu.Text = $"{versetCarteCurenta} {versetCapitolCurent}";

            // ✅ Căutăm listele în care se află versetele din capitol
            var listeCuRef = GetListeCuReferinte_DB(versetCarteCurenta, versetCapitolCurent);
            bool areFavorite = listeCuRef.Count > 0;

            bool areComentarii = ExistaComentarii(versetCarteCurenta, versetCapitolCurent);

            // ✅ Steaua apare sau dispare
            runStea.Text = areFavorite ? "⭐" : "";

            // ✅ Comentariile apar sau dispar
            txtComentarii.Visibility = areComentarii ? Visibility.Visible : Visibility.Collapsed;

            // ✅ Animația trebuie să pornească dacă există ORICARE dintre ele
            var tbSteaAnim = (TextBlock)runStea.Parent;
            var sbAnim = (Storyboard)tbSteaAnim.Resources["PulseStar"];

            if (areFavorite || areComentarii)
                sbAnim.Begin();
            else
                sbAnim.Stop();

            // ✅ Construim tooltip-ul COMPLET (stea + comentarii)
            var stack = new StackPanel();

            // ✅ Secțiunea stelei
            if (areFavorite)
            {
                var tbSteaTip = new TextBlock
                {
                    TextWrapping = TextWrapping.Wrap,
                    Width = 400,
                    Padding = new Thickness(12),
                    TextAlignment = System.Windows.TextAlignment.Left
                };

                tbSteaTip.Inlines.Add(new WpfRun("Titlu listă: ")
                {
                    Foreground = Brushes.DarkRed,
                    FontSize = 18
                });

                tbSteaTip.Inlines.Add(new WpfRun(string.Join(" | ", listeCuRef.Keys) + "\n")
                {
                    Foreground = Brushes.DarkBlue,
                    FontSize = 18,
                    FontWeight = FontWeights.Bold
                });

                tbSteaTip.Inlines.Add(new WpfRun($"{versetCarteCurenta} {versetCapitolCurent}\n")
                {
                    Foreground = Brushes.DarkGreen,
                    FontSize = 18,
                    FontWeight = FontWeights.Bold
                });

                tbSteaTip.Inlines.Add(new WpfRun("Salvat în liste:\n")
                {
                    Foreground = Brushes.DarkRed,
                    FontSize = 18
                });

                tbSteaTip.Inlines.Add(new WpfRun(string.Join("\n",
                    listeCuRef.Select(kvp => $"• {kvp.Key} — {string.Join(", ", kvp.Value)}")))
                {
                    Foreground = Brushes.Black,
                    FontSize = 18
                });

                stack.Children.Add(tbSteaTip);
            }

            // ✅ Secțiunea comentariilor
            if (areComentarii)
            {
                if (stack.Children.Count > 0)
                    stack.Children.Add(new Separator { Margin = new Thickness(0, 10, 0, 10) });

                var tbComent = new TextBlock
                {
                    TextWrapping = TextWrapping.Wrap,
                    Width = 350,
                    Padding = new Thickness(12),
                    TextAlignment = System.Windows.TextAlignment.Left
                };

                tbComent.Inlines.Add(new WpfRun("Comentarii disponibile pentru:\n")
                {
                    Foreground = Brushes.DarkRed,
                    FontSize = 18
                });

                tbComent.Inlines.Add(new WpfRun($"{versetCarteCurenta} {versetCapitolCurent}\n\n")
                {
                    Foreground = Brushes.DarkBlue,
                    FontSize = 18,
                    FontWeight = FontWeights.Bold
                });

                tbComent.Inlines.Add(new WpfRun("Apasă pe comentarii jos în bară și apoi pe semn → 📘 pentru a vedea comentariile.")
                {
                    Foreground = Brushes.Black,
                    FontSize = 16
                });

                stack.Children.Add(tbComent);
            }

            // ✅ Aplicăm tooltip-ul final
            txtStea.ToolTip = new ToolTip
            {
                Content = new ScrollViewer
                {
                    Content = stack,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    MaxHeight = 300
                }
            };
        }










        private List<string> GetListeInCareEsteVersetul(string carte, int capitol, int verset)
        {
            var liste = new List<string>();

            foreach (var kvp in favoritePeListe)
            {
                string numeLista = kvp.Key;
                var versete = kvp.Value;

                if (versete.Any(v =>
                    v.Carte == carte &&
                    v.Capitol == capitol &&
                    v.Verset == verset))
                {
                    liste.Add(numeLista);
                }
            }

            return liste;
        }





        private List<int> GetFavoriteVerseteDinDB(string carte, int capitol)
        {
            var rezultat = new List<int>();

            var cmd = new SQLiteCommand(@"
        SELECT Verset
        FROM FavoriteVersete
        WHERE Carte=@c AND Capitol=@cap
        ORDER BY Verset", conn);

            cmd.Parameters.AddWithValue("@c", carte);
            cmd.Parameters.AddWithValue("@cap", capitol);

            var reader = cmd.ExecuteReader();
            while (reader.Read())
                rezultat.Add(reader.GetInt32(0));

            reader.Close();
            return rezultat;
        }




        public void AdaugaLaSchita(VersetSchita verset)
        {
            if (string.IsNullOrWhiteSpace(verset.Carte) || verset.Capitol <= 0 || verset.Verset <= 0)
            {
                MessageBox.Show("VersetSchita incomplet — nu poate fi adăugat.");
                return;
            }

            if (schitaCurenta.Any(v =>
                v.Carte == verset.Carte &&
                v.Capitol == verset.Capitol &&
                v.Verset == verset.Verset))
            {
                MessageBox.Show("Versetul este deja în schiță.");
                return;
            }

            schitaCurenta.Add(verset);
            AfiseazaSchita();
        }


        private void btnArataStele_Click(object sender, RoutedEventArgs e)
        {
            // 🔥 Toggle pentru afișarea stelelor
            arataStele = !arataStele;

            // 🔥 Reîncarcă Scriptura cu stelele vizibile sau ascunse
            if (!string.IsNullOrEmpty(versetCarteCurenta) && versetCapitolCurent > 0)
                IncarcaVersete(versetCarteCurenta, versetCapitolCurent);

            // 🔥 ToolTip stilizat în funcție de stare
            btnArataStele.ToolTip = new ToolTip
            {
                Background = Brushes.LightYellow,
                Padding = new Thickness(6),
                Content = new TextBlock
                {
                    Text = arataStele ? "Ascunde butoanele ⭐" : "Arată butoanele ⭐",
                    Foreground = arataStele ? Brushes.Green : Brushes.DarkBlue,
                    FontSize = 12,
                    FontWeight = FontWeights.Bold
                }
            };

            ToolTipService.SetInitialShowDelay(btnArataStele, 0);
            ToolTipService.SetShowDuration(btnArataStele, 5000);
            ToolTipService.SetBetweenShowDelay(btnArataStele, 0);
        }







        //Favorite cu baza


        public Dictionary<string, List<VersetSchita>> IncarcaFavoriteDinBazaDeDate()
        {
            var rezultat = new Dictionary<string, List<VersetSchita>>();

            // 1) Încarcă listele
            var cmdListe = new SQLiteCommand("SELECT ListaId, Nume FROM ListeFavorite ORDER BY Nume", conn);
            var readerListe = cmdListe.ExecuteReader();

            var mapIdToName = new Dictionary<int, string>();

            while (readerListe.Read())
            {
                int id = readerListe.GetInt32(0);
                string nume = readerListe.GetString(1);

                mapIdToName[id] = nume;
                rezultat[nume] = new List<VersetSchita>();
            }
            readerListe.Close();

            // 2) Încarcă versetele din fiecare listă
            var cmdVersete = new SQLiteCommand(@"
        SELECT ListaId, Carte, Capitol, Verset, Text, Comentariu
        FROM FavoriteVersete
        ORDER BY ListaId, Capitol, Verset", conn);

            var readerVersete = cmdVersete.ExecuteReader();

            while (readerVersete.Read())
            {
                int listaId = readerVersete.GetInt32(0);
                string carte = readerVersete.GetString(1);
                int capitol = readerVersete.GetInt32(2);
                int verset = readerVersete.GetInt32(3);
                string text = readerVersete.GetString(4);
                string comentariu = readerVersete.IsDBNull(5) ? "" : readerVersete.GetString(5);

                if (mapIdToName.ContainsKey(listaId))
                {
                    rezultat[mapIdToName[listaId]].Add(new VersetSchita
                    {
                        Carte = carte,
                        Capitol = capitol,
                        Verset = verset,
                        Text = text,
                        Comentariu = comentariu
                    });
                }
            }

            readerVersete.Close();
            return rezultat;
        }
        private int AsiguraLista(string numeLista)
        {
            // verificăm dacă există
            var cmdCheck = new SQLiteCommand("SELECT ListaId FROM ListeFavorite WHERE Nume=@n", conn);
            cmdCheck.Parameters.AddWithValue("@n", numeLista);
            var result = cmdCheck.ExecuteScalar();

            if (result != null)
                return Convert.ToInt32(result);

            // dacă nu există → o creăm
            var cmdInsert = new SQLiteCommand("INSERT INTO ListeFavorite (Nume) VALUES (@n)", conn);
            cmdInsert.Parameters.AddWithValue("@n", numeLista);
            cmdInsert.ExecuteNonQuery();

            return (int)conn.LastInsertRowId;
        }

        private void AdaugaVersetLaLista(string numeLista, VersetSchita v)
        {
            int listaId = AsiguraLista(numeLista);

            // verificăm dacă există deja
            var cmdCheck = new SQLiteCommand(@"
        SELECT COUNT(*) FROM FavoriteVersete
        WHERE ListaId=@lid AND Carte=@c AND Capitol=@cap AND Verset=@v", conn);

            cmdCheck.Parameters.AddWithValue("@lid", listaId);
            cmdCheck.Parameters.AddWithValue("@c", v.Carte);
            cmdCheck.Parameters.AddWithValue("@cap", v.Capitol);
            cmdCheck.Parameters.AddWithValue("@v", v.Verset);

            if (Convert.ToInt32(cmdCheck.ExecuteScalar()) > 0)
                return;

            // inserăm
            var cmdInsert = new SQLiteCommand(@"
        INSERT INTO FavoriteVersete (ListaId, Carte, Capitol, Verset, Text, Comentariu)
        VALUES (@lid, @c, @cap, @v, @t, @com)", conn);

            cmdInsert.Parameters.AddWithValue("@lid", listaId);
            cmdInsert.Parameters.AddWithValue("@c", v.Carte);
            cmdInsert.Parameters.AddWithValue("@cap", v.Capitol);
            cmdInsert.Parameters.AddWithValue("@v", v.Verset);
            cmdInsert.Parameters.AddWithValue("@t", v.Text);
            cmdInsert.Parameters.AddWithValue("@com", v.Comentariu ?? "");

            cmdInsert.ExecuteNonQuery();
        }

        public void StergeLista(string numeLista)
        {
            var cmdGet = new SQLiteCommand("SELECT ListaId FROM ListeFavorite WHERE Nume=@n", conn);
            cmdGet.Parameters.AddWithValue("@n", numeLista);
            var result = cmdGet.ExecuteScalar();

            if (result == null)
                return;

            int listaId = Convert.ToInt32(result);

            // ștergem versetele
            var cmdDelV = new SQLiteCommand("DELETE FROM FavoriteVersete WHERE ListaId=@id", conn);
            cmdDelV.Parameters.AddWithValue("@id", listaId);
            cmdDelV.ExecuteNonQuery();

            // ștergem lista
            var cmdDelL = new SQLiteCommand("DELETE FROM ListeFavorite WHERE ListaId=@id", conn);
            cmdDelL.Parameters.AddWithValue("@id", listaId);
            cmdDelL.ExecuteNonQuery();
        }


        public void RedenumesteLista(string vechi, string nou)
        {
            var cmd = new SQLiteCommand("UPDATE ListeFavorite SET Nume=@nou WHERE Nume=@vechi", conn);
            cmd.Parameters.AddWithValue("@nou", nou);
            cmd.Parameters.AddWithValue("@vechi", vechi);
            cmd.ExecuteNonQuery();
        }

      

        private Dictionary<string, List<string>> GetListeCuReferinte_DB(string carte, int capitol)
        {
            var rezultat = new Dictionary<string, List<string>>();

            var cmd = new SQLiteCommand(@"
        SELECT L.Nume, F.Verset
        FROM FavoriteVersete F
        JOIN ListeFavorite L ON L.ListaId = F.ListaId
        WHERE F.Carte=@c AND F.Capitol=@cap
        ORDER BY F.Verset", conn);

            cmd.Parameters.AddWithValue("@c", carte);
            cmd.Parameters.AddWithValue("@cap", capitol);

            var reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                string lista = reader.GetString(0);
                int verset = reader.GetInt32(1);

                if (!rezultat.ContainsKey(lista))
                    rezultat[lista] = new List<string>();

                rezultat[lista].Add($"{carte} {capitol}:{verset}");
            }

            reader.Close();
            return rezultat;
        }

        public void StergeVersetDinLista(string lista, VersetSchita verset)
        {
            using var cmd = new SQLiteCommand(@"
        DELETE FROM FavoriteVersete
        WHERE ListaId = (SELECT ListaId FROM ListeFavorite WHERE Nume = @nume)
          AND Carte = @carte
          AND Capitol = @capitol
          AND Verset = @verset
    ", conn);

            cmd.Parameters.AddWithValue("@nume", lista);
            cmd.Parameters.AddWithValue("@carte", verset.Carte);
            cmd.Parameters.AddWithValue("@capitol", verset.Capitol);
            cmd.Parameters.AddWithValue("@verset", verset.Verset);

            cmd.ExecuteNonQuery();
        }


        public void ActualizeazaComentariu(string numeLista, VersetSchita verset)
        {
            var cmd = new SQLiteCommand(@"
        UPDATE FavoriteVersete
        SET Comentariu = @comentariu
        WHERE ListaId = (SELECT ListaId FROM ListeFavorite WHERE Nume = @lista)
          AND Carte = @carte
          AND Capitol = @capitol
          AND Verset = @verset
    ", conn);

            cmd.Parameters.AddWithValue("@comentariu", verset.Comentariu ?? "");
            cmd.Parameters.AddWithValue("@lista", numeLista);
            cmd.Parameters.AddWithValue("@carte", verset.Carte);
            cmd.Parameters.AddWithValue("@capitol", verset.Capitol);
            cmd.Parameters.AddWithValue("@verset", verset.Verset);

            cmd.ExecuteNonQuery();
        }

        //dublu clic pe versetu cautat si te duce la versetu biblic
        private void listVersete_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (listVersete.SelectedItem is not ListBoxItem item) return;
            if (item.Content is not TextBlock tb) return;
            if (tb.Inlines.FirstInline is not Run run) return;

            string titlu = run.Text.Trim();

            int idx = titlu.LastIndexOf(':');
            if (idx == -1) return;

            string versetStr = titlu.Substring(idx + 1).Split(' ')[0];
            if (!int.TryParse(versetStr, out int verset)) return;

            string before = titlu.Substring(0, idx).Trim();
            string capitolStr = before.Split(' ').Last();
            if (!int.TryParse(capitolStr, out int capitol)) return;

            string carte = before.Substring(0, before.Length - capitolStr.Length).Trim();

            // 🔥 AICI actualizăm labelul
            lblRezultateCautare.Visibility = Visibility.Visible;
            lblRezultateCautare.Content = $"{carte} {capitol}";

            // Navigăm
            NavigheazaLaVerset(carte, capitol, verset);
        }






        public void NavigheazaLaVerset(string carte, int capitol, int verset)
        {
            // actualizăm labelul de sus (folosim lblRezultateCautare)
            lblRezultateCautare.Visibility = Visibility.Visible;
            lblRezultateCautare.Content = $"{carte} {capitol}";

            // încărcăm capitolul
            IncarcaVerseteCapitol(carte, capitol);

            Dispatcher.InvokeAsync(() =>
            {
                foreach (ListBoxItem item in listVersete.Items)
                {
                    if (item.Content is TextBlock tb &&
                        tb.Inlines.FirstInline is Run run &&
                        run.Text.StartsWith($"{carte} {capitol}:{verset}"))
                    {
                        listVersete.SelectedItem = item;
                        item.BringIntoView();

                        EvidentiazaVerset(item);

                        break;
                    }
                }
            }, DispatcherPriority.Background);
        }







        public void IncarcaVerseteCapitol(string carte, int capitol)
        {
            listVersete.Items.Clear();

            var cmd = new SQLiteCommand(@"
SELECT Carte, Capitol, Verset, TextNormalizat 
FROM Versete 
WHERE Carte = @carte AND Capitol = @capitol
ORDER BY Verset", conn);

            cmd.Parameters.AddWithValue("@carte", carte);
            cmd.Parameters.AddWithValue("@capitol", capitol);

            var reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                string c = reader.GetString(0);
                int cap = reader.GetInt32(1);
                int vers = reader.GetInt32(2);
                string text = reader.GetString(3); // TextNormalizat cu <rosu>

                var tb = new TextBlock
                {
                    TextWrapping = TextWrapping.Wrap,
                    FontSize = marimeTextCurenta,
                    FontFamily = new WpfFontFamily("Segoe UI"),
                    Foreground = culoareTextVerset
                };

                // Titlul versetului
                tb.Inlines.Add(new Run($"{c} {cap}:{vers} — ") { Foreground = culoareTextVerset });

                // Procesăm textul cu <rosu>
                string t = text;

                while (true)
                {
                    int start = t.IndexOf("<rosu>");
                    if (start == -1)
                    {
                        tb.Inlines.Add(new Run(t) { Foreground = culoareTextVerset });
                        break;
                    }

                    if (start > 0)
                    {
                        tb.Inlines.Add(new Run(t.Substring(0, start)) { Foreground = culoareTextVerset });
                    }

                    t = t.Substring(start + 6);

                    int end = t.IndexOf("</rosu>");
                    if (end == -1)
                    {
                        tb.Inlines.Add(new Run(t) { Foreground = Brushes.Red });
                        break;
                    }

                    string rosu = t.Substring(0, end);
                    tb.Inlines.Add(new Run(rosu) { Foreground = Brushes.Red });

                    t = t.Substring(end + 7);
                }

                listVersete.Items.Add(new ListBoxItem
                {
                    Content = tb,
                    Foreground = culoareTextVerset
                });
            }
        }

        private async void EvidentiazaVerset(ListBoxItem item)
        {
            if (item.Content is not TextBlock tb) return;

            var original = tb.Background;

            tb.Background = Brushes.Yellow;   // evidențiere pe text

                  // revine la normal
        }

        private void txtCautare_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }
}




