using System.Data.SQLite;
using System.Windows;
using System.Windows.Controls;

namespace Biblia
{
    public partial class TemeWindow : Window
    {
        public TemeWindow()
        {
            InitializeComponent();

            using var conn = new SQLiteConnection("Data Source=biblia.db");
            conn.Open();

            var cmd = new SQLiteCommand("SELECT Id, Nume FROM Teme ORDER BY Id", conn);
            using var reader = cmd.ExecuteReader();

            int index = 1;
            while (reader.Read())
            {
                int id = reader.GetInt32(0);
                string nume = reader.GetString(1);

                string simbol = nume.ToUpperInvariant() switch
                {
                    var n when n.Contains("ADEVĂR") => "📏",
                    var n when n.Contains("ASCULTARE") => "👂",
                    var n when n.Contains("BOTEZ") => "🕊",
                    var n when n.Contains("CREDIN") => "📖",
                    var n when n.Contains("DRAGOSTE") || n.Contains("IUBIRE") => "❤️",
                    var n when n.Contains("DUHUL") => "🔥",
                    var n when n.Contains("HAR") => "💧",
                    var n when n.Contains("IERTARE") => "🧼",
                    var n when n.Contains("ÎMPĂRĂȚIA") => "👑",
                    var n when n.Contains("ÎNȚELEP") => "🧠",
                    var n when n.Contains("JERTF") => "🩸",
                    var n when n.Contains("JUDEC") => "⚖️",
                    var n when n.Contains("LUMIN") => "💡",
                    var n when n.Contains("MESIA") => "🔮",
                    var n when n.Contains("MILĂ") => "🤲",
                    var n when n.Contains("MÂNTUIR") => "🚪",
                    var n when n.Contains("MULȚUMIRE") => "🎁",
                    var n when n.Contains("NĂDEJDE") => "🪁",
                    var n when n.Contains("NEPRIHĂN") => "🧣",
                    var n when n.Contains("PĂCAT") => "⚠️",
                    var n when n.Contains("RUGĂCIUNE") => "🙏",
                    var n when n.Contains("SFINȚ") => "👼",
                    var n when n.Contains("VEȘNIC") => "⏳",
                    _ => "📘"
                };



                listTeme.Items.Add(new TemaItem(id, $"  {index}.  {simbol}  {nume}"));
                index++; 
            }
        }

       
        private TrimiteriWindow trimiteriWin;

        private void listTeme_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (listTeme.SelectedItem is TemaItem tema)
            {
                if (trimiteriWin == null || !trimiteriWin.IsLoaded)
                {
                    trimiteriWin = new TrimiteriWindow(tema);
                    trimiteriWin.Show();
                }
                else
                {
                    trimiteriWin.LoadTema(tema); // metodă nouă în TrimiteriWindow
                    trimiteriWin.Activate();
                }
            }
        }

    }
}





