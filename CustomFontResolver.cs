using PdfSharp.Fonts;
using System.IO;
using System.Reflection;

public class CustomFontResolver : IFontResolver
{
    public byte[] GetFont(string faceName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var resourceName = "Biblia.Fonts.LiberationSans-Regular.ttf"; // ← exact cum apare în listă

        using var stream = assembly.GetManifestResourceStream(resourceName);
        if (stream == null)
            throw new InvalidOperationException($"Font resource '{resourceName}' not found.");

        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms.ToArray();
    }

    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        return new FontResolverInfo("LiberationSans");
    }
}


