using System.Globalization;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace OpenChestnut.Codec;

// Centralizes bounded WordprocessingML revision discovery. Package operations
// use the same element catalog for scope checks, native-ID allocation, and
// post-write proof so revision policy cannot drift between entry points.
internal static class DocxRevisionMarkup
{
    internal static readonly HashSet<string> ElementNames = new(StringComparer.Ordinal)
    {
        "ins", "del",
        "moveFrom", "moveTo", "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
        "customXmlInsRangeStart", "customXmlInsRangeEnd", "customXmlDelRangeStart", "customXmlDelRangeEnd",
        "customXmlMoveFromRangeStart", "customXmlMoveFromRangeEnd", "customXmlMoveToRangeStart", "customXmlMoveToRangeEnd",
        "rPrChange", "pPrChange", "tblPrChange", "tblGridChange", "trPrChange", "tcPrChange", "sectPrChange", "numPrChange",
        "cellIns", "cellDel", "cellMerge",
    };

    internal static Dictionary<string, int> Inventory(byte[] bytes)
    {
        var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        VisitWordXml(bytes, (path, document) =>
        {
            result[path] = document.Descendants().Count(IsRevisionElement);
        });
        return result;
    }

    internal static HashSet<long> NativeIds(byte[] bytes)
    {
        var result = new HashSet<long>();
        VisitWordXml(bytes, (path, document) =>
        {
            foreach (var element in document.Descendants().Where(IsRevisionElement))
            {
                var attribute = element.Attributes().SingleOrDefault(value =>
                    value.Name.LocalName == "id" && IsWordprocessingNamespace(value.Name.NamespaceName));
                if (attribute is null) continue;
                if (!long.TryParse(attribute.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var value) || value < 0)
                    throw new CodecException(
                        "unsupported_document_revision_id",
                        $"DOCX revision element {element.Name.LocalName} has a non-decimal native ID.",
                        path);
                result.Add(value);
            }
        });
        return result;
    }

    internal static (string First, string Second) AllocatePair(byte[] bytes, ulong maxItems)
    {
        var inventory = Inventory(bytes);
        var count = inventory.Values.Aggregate(0UL, (sum, value) => checked(sum + (ulong)value));
        if (count > maxItems)
            throw new CodecException(
                "document_item_budget_exceeded",
                $"DOCX contains {count} revision markers and exceeds max_cells ({maxItems}).");

        var used = NativeIds(bytes);
        long candidate = 1;
        string Next()
        {
            while (used.Contains(candidate))
            {
                if (candidate == int.MaxValue)
                    throw new CodecException("document_revision_id_exhausted", "DOCX exhausted the positive 32-bit revision ID range.");
                candidate++;
            }
            if (candidate > int.MaxValue)
                throw new CodecException("document_revision_id_exhausted", "DOCX exhausted the positive 32-bit revision ID range.");
            var value = candidate.ToString(CultureInfo.InvariantCulture);
            used.Add(candidate);
            candidate++;
            return value;
        }

        return (Next(), Next());
    }

    internal static bool IsWordprocessingNamespace(string value) =>
        value.Contains("wordprocessingml", StringComparison.OrdinalIgnoreCase);

    private static bool IsRevisionElement(XElement element) =>
        IsWordprocessingNamespace(element.Name.NamespaceName) && ElementNames.Contains(element.Name.LocalName);

    private static void VisitWordXml(byte[] bytes, Action<string, XDocument> visitor)
    {
        try
        {
            using var stream = new MemoryStream(bytes, writable: false);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            foreach (var entry in archive.Entries.Where(entry =>
                         entry.FullName.StartsWith("word/", StringComparison.OrdinalIgnoreCase) &&
                         entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
            {
                using var partStream = entry.Open();
                using var reader = XmlReader.Create(partStream, new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                });
                visitor(entry.FullName, XDocument.Load(reader, LoadOptions.None));
            }
        }
        catch (XmlException exception)
        {
            throw new CodecException(
                "invalid_document_revision_xml",
                "DOCX contains malformed WordprocessingML while scanning revision scope.",
                innerException: exception);
        }
    }
}
