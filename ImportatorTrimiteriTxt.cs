using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;

public class ImportatorTrimiteriTxt
{
    public static void Import(string caleFisier)
    {
        using var conn = new SQLiteConnection("Data Source=biblia.db");
        conn.Open();

        using var transaction = conn.BeginTransaction();
        using var cmd = new SQLiteCommand("INSERT INTO Trimiteri (IdTrimitere, IdVerset, Tip, Referinta, Text) VALUES (@id, @verset, @tip, @ref, @text)", conn);
        cmd.Parameters.Add(new SQLiteParameter("@id"));
        cmd.Parameters.Add(new SQLiteParameter("@verset"));
        cmd.Parameters.Add(new SQLiteParameter("@tip"));
        cmd.Parameters.Add(new SQLiteParameter("@ref"));
        cmd.Parameters.Add(new SQLiteParameter("@text"));

        foreach (var line in File.ReadLines(caleFisier))
        {
            if (string.IsNullOrWhiteSpace(line) || line.StartsWith("IdTrimitere")) continue;

            var parts = SplitCSVLine(line);
            if (parts.Length < 5) continue;

            cmd.Parameters["@id"].Value = int.Parse(parts[0]);
            cmd.Parameters["@verset"].Value = int.Parse(parts[1]);
            cmd.Parameters["@tip"].Value = parts[2].Trim();
            cmd.Parameters["@ref"].Value = parts[3].Trim();
            cmd.Parameters["@text"].Value = parts[4].Trim();

            cmd.ExecuteNonQuery();
        }

        transaction.Commit();
        Console.WriteLine("✅ Importul din fișierul TXT s-a încheiat cu succes.");
    }

    private static string[] SplitCSVLine(string line)
    {
        var result = new List<string>();
        bool inQuotes = false;
        var current = "";

        foreach (char c in line)
        {
            if (c == '"') inQuotes = !inQuotes;
            else if (c == ',' && !inQuotes)
            {
                result.Add(current);
                current = "";
            }
            else current += c;
        }
        result.Add(current);
        return result.ToArray();
    }
}

