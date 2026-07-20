namespace OpenChestnut.Codec;

// ISO 29500 SpreadsheetML ST_IconSetType values from the base namespace.
// Office 2010 x14-only sets and custom icon graphs intentionally stay outside
// this catalog so every consumer shares the same fail-closed boundary.
internal static class XlsxIconSetCatalog
{
    private static readonly IReadOnlyDictionary<string, uint> Counts = new Dictionary<string, uint>(StringComparer.Ordinal)
    {
        ["3Arrows"] = 3, ["3ArrowsGray"] = 3, ["3Flags"] = 3, ["3TrafficLights1"] = 3,
        ["3TrafficLights2"] = 3, ["3Signs"] = 3, ["3Symbols"] = 3, ["3Symbols2"] = 3,
        ["4Arrows"] = 4, ["4ArrowsGray"] = 4, ["4RedToBlack"] = 4, ["4Rating"] = 4,
        ["4TrafficLights"] = 4, ["5Arrows"] = 5, ["5ArrowsGray"] = 5, ["5Rating"] = 5,
        ["5Quarters"] = 5,
    };

    internal static bool Contains(string value) => Counts.ContainsKey(value);

    internal static bool TryGetCount(string value, out uint count) => Counts.TryGetValue(value, out count);
}
