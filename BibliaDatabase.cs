using System;
using System.Data.SQLite;
using System.IO;

public static class BibliaDatabase
{
    public static void Initialize()
    {
        var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "biblia.db");

        // 🔹 Creează fișierul dacă nu există
        if (!File.Exists(path))
        {
            SQLiteConnection.CreateFile(path);
        }

        // 🔹 Deschide conexiunea
        using var conn = new SQLiteConnection($"Data Source={path}");
        conn.Open();

        // 🔹 Creează tabela Versete
        var cmdVersete = new SQLiteCommand(@"
            CREATE TABLE IF NOT EXISTS Versete (
                Carte TEXT NOT NULL,
                Capitol INTEGER NOT NULL,
                Verset INTEGER NOT NULL,
                Text TEXT NOT NULL,
                PRIMARY KEY (Carte, Capitol, Verset)
            );
        ", conn);
        cmdVersete.ExecuteNonQuery();

        // 🔹 Creează tabela Teme
        var cmdTeme = new SQLiteCommand(@"
            CREATE TABLE IF NOT EXISTS Teme (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Nume TEXT NOT NULL
            );
        ", conn);
        cmdTeme.ExecuteNonQuery();

        // 🔹 Creează tabela Trimiteri — AICI e locul corect
        var cmdTrimiteri = new SQLiteCommand(@"
            CREATE TABLE IF NOT EXISTS Trimiteri (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                TemaId INTEGER NOT NULL,
                Carte TEXT NOT NULL,
                Capitol INTEGER NOT NULL,
                Verset INTEGER NOT NULL,
                FOREIGN KEY (TemaId) REFERENCES Teme(Id)
            );
        ", conn);
        cmdTrimiteri.ExecuteNonQuery();

        // 🔹 Populează temele dacă tabela e goală
       
        InsertTeme();
        InsertTrimiteri();

        // 🔹 Confirmare disciplinată
        Console.WriteLine("✅ Initialize complet — tabele create.");
    }

    public static void InsertTeme()
    {
        using var conn = new SQLiteConnection("Data Source=biblia.db");
        conn.Open();

        var cmdCheck = new SQLiteCommand("SELECT COUNT(*) FROM Teme", conn);
        var count = Convert.ToInt32(cmdCheck.ExecuteScalar());

        if (count > 0) return;

        var teme = new[]
        {
            "Credința în Vechiul Testament",
            "Harul în epistolele lui Pavel",
            "Profeții despre Mesia",
            "Rugăciunea în Psalmi",
            "Împărăția lui Dumnezeu în Evanghelii"
        };

        foreach (var tema in teme)
        {
            var cmdInsert = new SQLiteCommand("INSERT INTO Teme (Nume) VALUES (@nume)", conn);
            cmdInsert.Parameters.AddWithValue("@nume", tema);
            cmdInsert.ExecuteNonQuery();
        }



        Console.WriteLine("✅ Teme inserate în baza de date.");
    }

    public static void InsertTrimiteri()
    {
        using var conn = new SQLiteConnection("Data Source=biblia.db");
        conn.Open();

        // 🔹 Verifică dacă tabela Trimiteri e deja populată
        var cmdCheck = new SQLiteCommand("SELECT COUNT(*) FROM Trimiteri", conn);
        var count = Convert.ToInt32(cmdCheck.ExecuteScalar());
        if (count > 0) return;

        // 🔹 Inseră versete pentru fiecare temă
        var trimiteri = new[]
        {
        // TemaId 1 — Credința în Vechiul Testament
        (1, "Evrei", 11, 1),
        (1, "Geneza", 15, 6),

        // TemaId 2 — Harul în epistolele lui Pavel
        (2, "Efeseni", 2, 8),
        (2, "Romani", 3, 24),

        // TemaId 3 — Profeții despre Mesia
        (3, "Isaia", 53, 5),
        (3, "Mica", 5, 2),

        // TemaId 4 — Rugăciunea în Psalmi
        (4, "Psalmii", 23, 1),
        (4, "Psalmii", 51, 10),

        // TemaId 5 — Împărăția lui Dumnezeu în Evanghelii
        (5, "Matei", 6, 33),
        (5, "Luca", 17, 21)
    };

        foreach (var (temaId, carte, capitol, verset) in trimiteri)
        {
            var cmdInsert = new SQLiteCommand(@"
            INSERT INTO Trimiteri (TemaId, Carte, Capitol, Verset)
            VALUES (@temaId, @carte, @capitol, @verset);
        ", conn);

            cmdInsert.Parameters.AddWithValue("@temaId", temaId);
            cmdInsert.Parameters.AddWithValue("@carte", carte);
            cmdInsert.Parameters.AddWithValue("@capitol", capitol);
            cmdInsert.Parameters.AddWithValue("@verset", verset);
            cmdInsert.ExecuteNonQuery();
        }

        Console.WriteLine("✅ Trimiteri inserate în baza de date.");
    }

}


