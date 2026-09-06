using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

public static class BibliaImporter
{
    public static void ImportVersete(SQLiteConnection conn)
    {
        var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Versete.txt");
        if (!File.Exists(path)) return;

        var lines = File.ReadAllLines(path);
        for (int i = 1; i < lines.Length; i++) // ignoră headerul
        {
            var parts = SplitCsvLine(lines[i]);
            if (parts.Length < 7) continue;

            var carte = parts[4].Trim();
            var capitol = int.Parse(parts[5]);
            var verset = int.Parse(parts[6]);
            var text = parts[3].Trim();
            var carteOrdine = GetCarteOrdine(carte);

            var cmd = new SQLiteCommand(@"
                INSERT INTO Versete (Carte, Capitol, Verset, Text, CarteOrdine)
                VALUES (@carte, @capitol, @verset, @text, @ordine)", conn);

            cmd.Parameters.AddWithValue("@carte", carte);
            cmd.Parameters.AddWithValue("@capitol", capitol);
            cmd.Parameters.AddWithValue("@verset", verset);
            cmd.Parameters.AddWithValue("@text", text);
            cmd.Parameters.AddWithValue("@ordine", carteOrdine);
            cmd.ExecuteNonQuery();
        }
    }

    public static void ImportTrimiteriTXT(SQLiteConnection conn)
    {
        var path = @"C:\Users\Furdui\Desktop\Biblia\Biblia\Biblia\Trimiteri.txt";
        if (!File.Exists(path))
        {
            Console.WriteLine("❌ Fișierul Trimiteri.txt nu există.");
            return;
        }

        var lines = File.ReadAllLines(path);
        for (int i = 1; i < lines.Length; i++) // ignoră header
        {
            var parts = lines[i].Split(',');
            if (parts.Length < 4) continue;

            var idVerset = int.Parse(parts[1]);
            var tip = parts[2].Trim();
            var referintaRaw = parts[3].Trim().Trim('"').Replace("*", "").Trim();

            var grupuri = referintaRaw.Split(';');
            foreach (var grup in grupuri)
            {
                var bloc = grup.Trim();
                if (string.IsNullOrEmpty(bloc)) continue;

                var match = Regex.Match(bloc, @"^([^\d]+)(\d+):([\d,]+)$");
                if (!match.Success)
                {
                    Inserare(conn, idVerset, tip, bloc);
                    continue;
                }

                var carte = match.Groups[1].Value.Trim();
                var capitol = match.Groups[2].Value.Trim();
                var versete = match.Groups[3].Value.Split(',');

                foreach (var v in versete)
                {
                    var referinta = $"{carte}{capitol}:{v.Trim()}";
                    Inserare(conn, idVerset, tip, referinta);
                }
            }
        }

        Console.WriteLine("✅ Import din Trimiteri.txt finalizat.");
    }


    private static void Inserare(SQLiteConnection conn, int idVerset, string tip, string referinta)
    {
        var cmdCheck = new SQLiteCommand(@"
        SELECT COUNT(*) FROM TReferinta 
        WHERE IdVerset = @idVerset AND Tip = @tip AND Referinta = @referinta", conn);
        cmdCheck.Parameters.AddWithValue("@idVerset", idVerset);
        cmdCheck.Parameters.AddWithValue("@tip", tip);
        cmdCheck.Parameters.AddWithValue("@referinta", referinta);

        var exists = Convert.ToInt32(cmdCheck.ExecuteScalar());
        if (exists > 0) return;

        var cmdInsert = new SQLiteCommand(@"
        INSERT INTO TReferinta (IdVerset, Tip, Referinta)
        VALUES (@idVerset, @tip, @referinta)", conn);
        cmdInsert.Parameters.AddWithValue("@idVerset", idVerset);
        cmdInsert.Parameters.AddWithValue("@tip", tip);
        cmdInsert.Parameters.AddWithValue("@referinta", referinta);
        cmdInsert.ExecuteNonQuery();
    }




    private static int GetCarteOrdine(string carte)
    {
        var canon = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["GENESA1"] = 1,
            ["EXODUL"] = 2,
            ["LEVITICUL"] = 3,
            ["NUMERI"] = 4,
            ["DEUTERONOM"] = 5,
            ["IOSUA"] = 6,
            ["JUDECĂTORI"] = 7,
            ["RUT"] = 8,
            ["1 SAMUEL"] = 9,
            ["2 SAMUEL"] = 10,
            ["1 ÎMPĂRAȚI"] = 11,
            ["2 ÎMPĂRAȚI"] = 12,
            ["1 CRONICI"] = 13,
            ["2 CRONICI"] = 14,
            ["EZRA"] = 15,
            ["NEEMIA"] = 16,
            ["ESTERA"] = 17,
            ["IOB"] = 18,
            ["PSALMII"] = 19,
            ["PROVERBE"] = 20,
            ["ECLESIASTUL"] = 21,
            ["CÂNTAREA CÂNTĂRILOR"] = 22,
            ["ISAIA"] = 23,
            ["IEREMIA"] = 24,
            ["PLANGERILE LUI EREMIA"] = 25,
            ["EZECHEL"] = 26,
            ["DANIEL"] = 27,
            ["OSEA"] = 28,
            ["IOEL"] = 29,
            ["AMOS"] = 30,
            ["OBADIA"] = 31,
            ["IONA"] = 32,
            ["MIICA"] = 33,
            ["NAUM"] = 34,
            ["HABACUC"] = 35,
            ["ȚEFANIA"] = 36,
            ["HAGAI"] = 37,
            ["ZAHARIA"] = 38,
            ["MALEAHI"] = 39,
            ["MATEI"] = 40,
            ["MARCU"] = 41,
            ["LUCA"] = 42,
            ["IOAN"] = 43,
            ["FAPTE"] = 44,
            ["ROMANI"] = 45,
            ["1 CORINTENI"] = 46,
            ["2 CORINTENI"] = 47,
            ["GALATENI"] = 48,
            ["EFESENI"] = 49,
            ["FILIPENI"] = 50,
            ["COLOSENI"] = 51,
            ["1 TESALONICENI"] = 52,
            ["2 TESALONICENI"] = 53,
            ["1 TIMOTEI"] = 54,
            ["2 TIMOTEI"] = 55,
            ["TIT"] = 56,
            ["FILIMON"] = 57,
            ["EVREI"] = 58,
            ["IACOV"] = 59,
            ["1 PETRU"] = 60,
            ["2 PETRU"] = 61,
            ["1 IOAN"] = 62,
            ["2 IOAN"] = 63,
            ["3 IOAN"] = 64,
            ["IUDA"] = 65,
            ["APOCALIPSA"] = 66
        };

        return canon.TryGetValue(carte.ToUpperInvariant(), out var ordine) ? ordine : 0;
    }

    private static string[] SplitCsvLine(string line)
    {
        var values = new List<string>();
        bool inQuotes = false;
        var current = new StringBuilder();

        foreach (char c in line)
        {
            if (c == '\"') inQuotes = !inQuotes;
            else if (c == ',' && !inQuotes)
            {
                values.Add(current.ToString());
                current.Clear();
            }
            else current.Append(c);
        }

        values.Add(current.ToString());
        return values.ToArray();
    }
}


