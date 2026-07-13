using A = DocumentFormat.OpenXml.Drawing;

namespace OpenOffice.OpenXmlCodec;

internal static class PptxColor
{
    private static readonly IReadOnlyDictionary<string, A.SchemeColorValues> SchemeColors =
        new Dictionary<string, A.SchemeColorValues>(StringComparer.Ordinal)
        {
            ["bg1"] = A.SchemeColorValues.Background1,
            ["tx1"] = A.SchemeColorValues.Text1,
            ["bg2"] = A.SchemeColorValues.Background2,
            ["tx2"] = A.SchemeColorValues.Text2,
            ["accent1"] = A.SchemeColorValues.Accent1,
            ["accent2"] = A.SchemeColorValues.Accent2,
            ["accent3"] = A.SchemeColorValues.Accent3,
            ["accent4"] = A.SchemeColorValues.Accent4,
            ["accent5"] = A.SchemeColorValues.Accent5,
            ["accent6"] = A.SchemeColorValues.Accent6,
            ["hlink"] = A.SchemeColorValues.Hyperlink,
            ["folHlink"] = A.SchemeColorValues.FollowedHyperlink,
            ["dk1"] = A.SchemeColorValues.Dark1,
            ["lt1"] = A.SchemeColorValues.Light1,
            ["dk2"] = A.SchemeColorValues.Dark2,
            ["lt2"] = A.SchemeColorValues.Light2,
        };

    internal static string SolidRgb(A.SolidFill? fill) =>
        fill?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value ?? string.Empty;

    internal static string Normalize(string value)
    {
        var rgb = value.Trim().TrimStart('#').ToUpperInvariant();
        if (rgb.Length != 6 || rgb.Any(character => !Uri.IsHexDigit(character)))
            throw new CodecException("invalid_presentation_color", $"Presentation color {value} must be a six-digit RGB value.");
        return rgb;
    }

    internal static string NormalizeScheme(string value)
    {
        var scheme = value.Trim();
        if (!SchemeColors.ContainsKey(scheme))
            throw new CodecException("invalid_presentation_color", $"Presentation scheme color {value} is not a supported theme token.");
        return scheme;
    }

    internal static A.SchemeColorValues SchemeValue(string value) => SchemeColors[NormalizeScheme(value)];

    internal static bool TrySchemeToken(A.SchemeColorValues value, out string token)
    {
        foreach (var entry in SchemeColors)
        {
            if (!entry.Value.Equals(value)) continue;
            token = entry.Key;
            return true;
        }
        token = string.Empty;
        return false;
    }
}
