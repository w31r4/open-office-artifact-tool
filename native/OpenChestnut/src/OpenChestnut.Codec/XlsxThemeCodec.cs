using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns only the public 12-slot Spreadsheet theme projection. Font schemes,
// format schemes, color transforms, extensions, and whitespace stay in the
// source ThemePart and are preserved unless the corresponding modeled color
// is explicitly replaced.
internal sealed class XlsxThemeCodec
{
    private static readonly XNamespace Drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly string[] SlotNames = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
    private static readonly string[] DefaultColors = ["000000", "FFFFFF", "1F497D", "EEECE1", "4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646", "0000FF", "800080"];

    private readonly WorkbookPart _workbookPart;
    private ThemePart? _part;
    private XDocument? _document;
    private SpreadsheetThemeArtifact? _current;
    private SpreadsheetThemeSourceBinding? _binding;
    private bool _editable;

    internal XlsxThemeCodec(WorkbookPart workbookPart)
    {
        _workbookPart = workbookPart;
        var themeParts = workbookPart.Parts.Where(item => item.OpenXmlPart is ThemePart).ToArray();
        if (themeParts.Length > 1) throw Invalid("Workbook contains more than one internal theme relationship.");
        _part = themeParts.Length == 1 ? themeParts[0].OpenXmlPart as ThemePart : null;
        if (_part is not null) LoadSourceTheme();
    }

    internal bool OwnsOpaqueTheme => _part is not null && Dirty;
    internal bool Dirty { get; private set; }
    internal string? PartPath => _part?.Uri.OriginalString.TrimStart('/');

    internal SpreadsheetThemeArtifact? Read() => _current?.Clone();

    internal bool OwnsPart(OpaqueOpcPart part) => OwnsOpaqueTheme && PartPath is { } path && part.Path.Equals(path, StringComparison.OrdinalIgnoreCase);

    internal static bool IsThemeRelationship(OpaqueOpcRelationship relationship) =>
        relationship.SourcePath.Equals("xl/workbook.xml", StringComparison.OrdinalIgnoreCase) &&
        relationship.TargetMode.Length == 0 &&
        relationship.Type.EndsWith("/theme", StringComparison.Ordinal);

    internal void Apply(SpreadsheetThemeArtifact? desired, bool sourceBound)
    {
        if (_part is not null && sourceBound) ValidateSourceBinding(desired?.Source);

        if (desired is null)
        {
            if (_part is not null) throw Invalid("Source-preserving XLSX export cannot remove the workbook theme.");
            return;
        }

        if (!HasCompleteSemantics(desired))
        {
            if (_part is not null && sourceBound && desired.Source is { Editable: false }) return;
            throw Invalid("Workbook theme must provide a name and all 12 six-digit RGB color slots.");
        }
        Validate(desired);

        if (_part is null)
        {
            if (sourceBound && IsDefaultTheme(desired)) return;
            _part = _workbookPart.AddNewPart<ThemePart>();
            _document = CreateThemeDocument(desired);
            _current = SemanticClone(desired);
            _editable = true;
            Dirty = true;
            return;
        }

        if (!_editable) throw Invalid("The source workbook theme uses color semantics outside the bounded sRGB/system-color profile and cannot be replaced losslessly.");
        if (SemanticSha256(desired).Equals(_binding?.SemanticSha256, StringComparison.OrdinalIgnoreCase)) return;

        var root = _document?.Root ?? throw Invalid("Workbook ThemePart has no theme root.");
        var scheme = root.Element(Drawing + "themeElements")?.Element(Drawing + "clrScheme") ?? throw Invalid("Workbook theme has no color scheme.");
        root.SetAttributeValue("name", desired.Name);
        var colors = ThemeColors(desired);
        var currentColors = ThemeColors(_current!);
        for (var index = 0; index < SlotNames.Length; index++)
        {
            if (colors[index].Equals(currentColors[index], StringComparison.OrdinalIgnoreCase)) continue;
            var slot = scheme.Elements(Drawing + SlotNames[index]).Single();
            slot.Elements().Single().ReplaceWith(new XElement(Drawing + "srgbClr", new XAttribute("val", colors[index])));
        }
        _current = SemanticClone(desired);
        Dirty = true;
    }

    internal void Save()
    {
        if (!Dirty || _part is null || _document is null) return;
        using var stream = _part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            OmitXmlDeclaration = false,
        });
        _document.Save(writer);
    }

    internal static void Validate(SpreadsheetThemeArtifact? theme)
    {
        if (theme is null) return;
        if (!HasCompleteSemantics(theme))
        {
            if (theme.Source is { Editable: false, PartPath.Length: > 0, XmlSha256.Length: > 0 }) return;
            throw Invalid("Workbook theme must provide a name and all 12 six-digit RGB color slots.");
        }
        if (theme.Name.Length > 255 || theme.Name.Any(char.IsControl)) throw Invalid("Workbook theme name must be at most 255 characters and contain no control characters.");
        foreach (var (slot, color) in SlotNames.Zip(ThemeColors(theme)))
            if (color.Length != 6 || !color.All(Uri.IsHexDigit)) throw Invalid($"Workbook theme slot {slot} must be six-digit RGB.");
    }

    private void LoadSourceTheme()
    {
        using var source = _part!.GetStream(FileMode.Open, FileAccess.Read);
        using var copy = new MemoryStream();
        source.CopyTo(copy);
        var bytes = copy.ToArray();
        try
        {
            using var reader = XmlReader.Create(new MemoryStream(bytes, writable: false), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            _document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }
        catch (XmlException exception)
        {
            throw new CodecException("invalid_workbook_theme", "Workbook ThemePart is not valid XML.", PartPath, exception);
        }

        _editable = TryReadSemantics(_document, out var semantic);
        _binding = new SpreadsheetThemeSourceBinding
        {
            PartPath = PartPath ?? string.Empty,
            XmlSha256 = Sha256(bytes),
            SemanticSha256 = _editable ? SemanticSha256(semantic!) : string.Empty,
            Editable = _editable,
        };
        _current = semantic ?? new SpreadsheetThemeArtifact();
        _current.Source = _binding.Clone();
    }

    private void ValidateSourceBinding(SpreadsheetThemeSourceBinding? binding)
    {
        if (binding is null || _binding is null) throw Invalid("Source-bound workbook theme is missing its source binding.");
        if (!binding.PartPath.Equals(_binding.PartPath, StringComparison.OrdinalIgnoreCase) ||
            !binding.XmlSha256.Equals(_binding.XmlSha256, StringComparison.OrdinalIgnoreCase) ||
            binding.Editable != _binding.Editable ||
            !binding.SemanticSha256.Equals(_binding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
            throw Invalid("Workbook theme source binding does not match the validated source package.");
    }

    private static bool TryReadSemantics(XDocument document, out SpreadsheetThemeArtifact? theme)
    {
        theme = null;
        var root = document.Root;
        var scheme = root?.Name == Drawing + "theme" ? root.Element(Drawing + "themeElements")?.Element(Drawing + "clrScheme") : null;
        if (root is null || scheme is null) return false;
        var colors = new string[SlotNames.Length];
        for (var index = 0; index < SlotNames.Length; index++)
        {
            var slots = scheme.Elements(Drawing + SlotNames[index]).ToArray();
            if (slots.Length != 1 || slots[0].Elements().Count() != 1) return false;
            var color = slots[0].Elements().Single();
            var candidate = color.Name == Drawing + "srgbClr"
                ? color.Attribute("val")?.Value
                : color.Name == Drawing + "sysClr" ? color.Attribute("lastClr")?.Value ?? color.Attribute("val")?.Value : null;
            if (candidate is null || candidate.Length != 6 || !candidate.All(Uri.IsHexDigit)) return false;
            colors[index] = candidate.ToUpperInvariant();
        }
        theme = FromColors(root.Attribute("name")?.Value is { Length: > 0 } name ? name : "Imported Office Theme", colors);
        return true;
    }

    private static XDocument CreateThemeDocument(SpreadsheetThemeArtifact theme)
    {
        var colors = ThemeColors(theme);
        XElement SolidFill() => new(Drawing + "solidFill", new XElement(Drawing + "schemeClr", new XAttribute("val", "phClr")));
        XElement Line(int width) => new(Drawing + "ln", new XAttribute("w", width), SolidFill());
        var colorScheme = new XElement(Drawing + "clrScheme", new XAttribute("name", theme.Name),
            SlotNames.Select((slot, index) => new XElement(Drawing + slot, new XElement(Drawing + "srgbClr", new XAttribute("val", colors[index])))));
        var fontScheme = new XElement(Drawing + "fontScheme", new XAttribute("name", "Office Clean Room"),
            new XElement(Drawing + "majorFont", new XElement(Drawing + "latin", new XAttribute("typeface", "Aptos Display")), new XElement(Drawing + "ea", new XAttribute("typeface", "")), new XElement(Drawing + "cs", new XAttribute("typeface", ""))),
            new XElement(Drawing + "minorFont", new XElement(Drawing + "latin", new XAttribute("typeface", "Aptos")), new XElement(Drawing + "ea", new XAttribute("typeface", "")), new XElement(Drawing + "cs", new XAttribute("typeface", ""))));
        var formatScheme = new XElement(Drawing + "fmtScheme", new XAttribute("name", "Office Clean Room"),
            new XElement(Drawing + "fillStyleLst", SolidFill(), SolidFill(), SolidFill()),
            new XElement(Drawing + "lnStyleLst", Line(9525), Line(25400), Line(38100)),
            new XElement(Drawing + "effectStyleLst",
                new XElement(Drawing + "effectStyle", new XElement(Drawing + "effectLst")),
                new XElement(Drawing + "effectStyle", new XElement(Drawing + "effectLst")),
                new XElement(Drawing + "effectStyle", new XElement(Drawing + "effectLst"))),
            new XElement(Drawing + "bgFillStyleLst", SolidFill(), SolidFill(), SolidFill()));
        return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(Drawing + "theme", new XAttribute(XNamespace.Xmlns + "a", Drawing), new XAttribute("name", theme.Name),
                new XElement(Drawing + "themeElements", colorScheme, fontScheme, formatScheme)));
    }

    private static SpreadsheetThemeArtifact SemanticClone(SpreadsheetThemeArtifact source) => FromColors(source.Name, ThemeColors(source));

    private static SpreadsheetThemeArtifact FromColors(string name, IReadOnlyList<string> colors) => new()
    {
        Name = name,
        Dk1Rgb = colors[0], Lt1Rgb = colors[1], Dk2Rgb = colors[2], Lt2Rgb = colors[3],
        Accent1Rgb = colors[4], Accent2Rgb = colors[5], Accent3Rgb = colors[6], Accent4Rgb = colors[7],
        Accent5Rgb = colors[8], Accent6Rgb = colors[9], HlinkRgb = colors[10], FolHlinkRgb = colors[11],
    };

    private static string[] ThemeColors(SpreadsheetThemeArtifact theme) =>
    [
        theme.Dk1Rgb, theme.Lt1Rgb, theme.Dk2Rgb, theme.Lt2Rgb,
        theme.Accent1Rgb, theme.Accent2Rgb, theme.Accent3Rgb, theme.Accent4Rgb,
        theme.Accent5Rgb, theme.Accent6Rgb, theme.HlinkRgb, theme.FolHlinkRgb,
    ];

    private static bool HasCompleteSemantics(SpreadsheetThemeArtifact theme) =>
        !string.IsNullOrWhiteSpace(theme.Name) && ThemeColors(theme).All(color => !string.IsNullOrWhiteSpace(color));

    private static bool IsDefaultTheme(SpreadsheetThemeArtifact theme) =>
        theme.Name == "Office Clean Room" && ThemeColors(theme).SequenceEqual(DefaultColors, StringComparer.OrdinalIgnoreCase);

    private static string SemanticSha256(SpreadsheetThemeArtifact theme) =>
        Sha256(Encoding.UTF8.GetBytes(string.Join('\0', [theme.Name, .. ThemeColors(theme).Select(color => color.ToUpperInvariant())])));

    private static string Sha256(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static CodecException Invalid(string message) => new("invalid_workbook_theme", message, "xl/theme/theme1.xml");
}
