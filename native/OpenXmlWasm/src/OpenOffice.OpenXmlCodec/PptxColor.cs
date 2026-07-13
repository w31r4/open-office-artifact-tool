using A = DocumentFormat.OpenXml.Drawing;

namespace OpenOffice.OpenXmlCodec;

internal static class PptxColor
{
    internal static string SolidRgb(A.SolidFill? fill) =>
        fill?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value ?? string.Empty;

    internal static string Normalize(string value)
    {
        var rgb = value.Trim().TrimStart('#').ToUpperInvariant();
        if (rgb.Length != 6 || rgb.Any(character => !Uri.IsHexDigit(character)))
            throw new CodecException("invalid_presentation_color", $"Presentation color {value} must be a six-digit RGB value.");
        return rgb;
    }
}
