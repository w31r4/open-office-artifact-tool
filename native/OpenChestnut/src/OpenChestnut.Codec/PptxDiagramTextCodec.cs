using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

internal sealed record PptxDiagramTextReplacement(string PartPath, string Sha256, byte[] Data);

// Owns one deliberately small SmartArt edit boundary. It does not author a
// diagram, change the graph, or reinterpret layout/style/colors. It exposes
// text only where an imported top-level p:graphicFrame proves it owns the
// canonical closed four-part Diagram graph and every document point has one
// direct plain DrawingML run. Everything outside that profile stays opaque.
internal static class PptxDiagramTextCodec
{
    private const string DiagramDataContentType = "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml";
    private const int MaxModelIdLength = 1_024;
    private const int MaxNodeTextLength = 32_767;

    private static readonly HashSet<string> DiagramNamespaces = new(StringComparer.Ordinal)
    {
        "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        "http://purl.oclc.org/ooxml/drawingml/diagram",
    };

    private static readonly HashSet<string> DrawingNamespaces = new(StringComparer.Ordinal)
    {
        "http://schemas.openxmlformats.org/drawingml/2006/main",
        "http://purl.oclc.org/ooxml/drawingml/main",
    };

    private sealed record DiagramNode(string ModelId, string Text, XElement TextElement);

    private sealed record ResolvedDiagram(
        PresentationDiagramText Binding,
        DiagramDataPart Part,
        XDocument Document,
        IReadOnlyList<DiagramNode> Nodes);

    internal static bool TryDescribe(OpenXmlElement source, OpenXmlPart owner, out PresentationDiagramText binding)
    {
        if (!TryResolve(source, owner, out var resolved))
        {
            binding = null!;
            return false;
        }
        binding = resolved.Binding;
        return true;
    }

    internal static PptxDiagramTextReplacement? PrepareReplacement(
        SlidePart owner,
        OpenXmlElement source,
        PresentationOpaqueElement original,
        PresentationOpaqueElement requested)
    {
        if (original.DiagramText is null)
        {
            if (requested.DiagramText is not null)
                throw Unsupported("An unrecognized SmartArt graph cannot claim the bounded diagram-text edit capability.");
            return null;
        }
        if (requested.DiagramText is null)
            throw Unsupported("A source-bound SmartArt text binding cannot be removed.");
        if (!TryResolve(source, owner, out var resolved))
            throw BindingMismatch("The SmartArt source no longer proves the bounded diagram-text profile.", PartPath(owner));
        if (!SameBinding(original.DiagramText, resolved.Binding) || !SameNodes(original.DiagramText.Nodes, resolved.Nodes))
            throw BindingMismatch("The SmartArt diagram data no longer matches its source binding.", resolved.Binding.PartPath);
        ValidateRequestedNodes(original.DiagramText.Nodes, requested.DiagramText.Nodes, resolved.Binding.PartPath);

        var changed = false;
        for (var index = 0; index < resolved.Nodes.Count; index++)
        {
            var requestedText = requested.DiagramText.Nodes[index].Text;
            if (resolved.Nodes[index].Text == requestedText) continue;
            SetText(resolved.Nodes[index].TextElement, requestedText);
            changed = true;
        }
        if (!changed) return null;

        var data = Serialize(resolved.Document);
        return new PptxDiagramTextReplacement(resolved.Binding.PartPath, Hash(data), data);
    }

    internal static void Apply(
        SlidePart owner,
        PresentationDiagramText binding,
        PptxDiagramTextReplacement replacement)
    {
        DiagramDataPart part;
        try
        {
            part = owner.GetPartById(binding.RelationshipId) as DiagramDataPart
                ?? throw BindingMismatch("The SmartArt data relationship no longer resolves to a DiagramDataPart.", PartPath(owner));
        }
        catch (ArgumentOutOfRangeException exception)
        {
            throw BindingMismatch("The SmartArt data relationship no longer resolves to a package part.", PartPath(owner), exception);
        }
        var partPath = PartPath(part);
        if (!partPath.Equals(binding.PartPath, StringComparison.OrdinalIgnoreCase) ||
            !partPath.Equals(replacement.PartPath, StringComparison.OrdinalIgnoreCase) ||
            !part.ContentType.Equals(binding.ContentType, StringComparison.OrdinalIgnoreCase) ||
            !part.ContentType.Equals(DiagramDataContentType, StringComparison.OrdinalIgnoreCase) ||
            !part.RelationshipType.EndsWith("/diagramData", StringComparison.Ordinal))
            throw BindingMismatch("The SmartArt data part path, content type, or relationship type no longer matches its source binding.", partPath);
        if (!Hash(ReadPart(part)).Equals(binding.SourceSha256, StringComparison.OrdinalIgnoreCase))
            throw BindingMismatch("The SmartArt data bytes no longer match their source digest.", partPath);

        using var output = part.GetStream(FileMode.Create, FileAccess.Write);
        output.Write(replacement.Data);
    }

    internal static void ValidateSourceBoundOutput(
        SlidePart sourceOwner,
        SlidePart outputOwner,
        OpenXmlElement source,
        OpenXmlElement output,
        PresentationOpaqueElement requested)
    {
        if (requested.DiagramText is null) return;
        if (!TryResolve(source, sourceOwner, out var sourceResolved) ||
            !SameBinding(requested.DiagramText, sourceResolved.Binding) ||
            !SameNodeIds(requested.DiagramText.Nodes, sourceResolved.Nodes))
            throw BindingMismatch("The source SmartArt diagram text does not match the requested source binding.", PartPath(sourceOwner));
        if (!TryResolve(output, outputOwner, out var outputResolved))
            throw BindingMismatch("The exported SmartArt diagram no longer proves the bounded diagram-text profile.", PartPath(outputOwner));
        if (!outputResolved.Binding.PartPath.Equals(requested.DiagramText.PartPath, StringComparison.OrdinalIgnoreCase) ||
            !outputResolved.Binding.ContentType.Equals(requested.DiagramText.ContentType, StringComparison.OrdinalIgnoreCase) ||
            outputResolved.Binding.RelationshipId != requested.DiagramText.RelationshipId ||
            !SameRequestedText(requested.DiagramText.Nodes, outputResolved.Nodes))
            throw BindingMismatch("The exported SmartArt node text does not match the requested bounded edit.", outputResolved.Binding.PartPath);
    }

    private static bool TryResolve(OpenXmlElement source, OpenXmlPart owner, out ResolvedDiagram resolved)
    {
        resolved = null!;
        if (source is not P.GraphicFrame frame || owner is not SlidePart || source.Parent is not P.ShapeTree ||
            PptxNativeObjectCatalog.Classify(source) != "diagram" || !PptxNativeObjectCatalog.SupportsPlacementEditing(source))
            return false;
        var roots = frame.Descendants().Where(PptxNativeObjectCatalog.IsDiagramRelationshipIds).ToArray();
        if (roots.Length != 1) return false;
        var relationshipAttributes = new[] { (OpenXmlElement)frame }.Concat(frame.Descendants())
            .SelectMany(element => element.GetAttributes())
            .Where(attribute => PptxNativeObjectCatalog.IsRelationshipNamespace(attribute.NamespaceUri))
            .ToArray();
        var rootAttributes = roots[0].GetAttributes()
            .Where(attribute => PptxNativeObjectCatalog.IsRelationshipNamespace(attribute.NamespaceUri))
            .ToArray();
        if (relationshipAttributes.Length != 4 || rootAttributes.Length != 4) return false;

        var expected = new Dictionary<string, Type>(StringComparer.Ordinal)
        {
            ["dm"] = typeof(DiagramDataPart),
            ["lo"] = typeof(DiagramLayoutDefinitionPart),
            ["qs"] = typeof(DiagramStylePart),
            ["cs"] = typeof(DiagramColorsPart),
        };
        var relationshipIds = new HashSet<string>(StringComparer.Ordinal);
        var partPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        DiagramDataPart? dataPart = null;
        string dataRelationshipId = string.Empty;
        foreach (var attribute in rootAttributes)
        {
            if (!expected.TryGetValue(attribute.LocalName, out var expectedType) || string.IsNullOrWhiteSpace(attribute.Value) ||
                !relationshipIds.Add(attribute.Value)) return false;
            OpenXmlPart part;
            try
            {
                part = owner.GetPartById(attribute.Value);
            }
            catch (ArgumentOutOfRangeException)
            {
                return false;
            }
            if (part.GetType() != expectedType || !IsClosedDiagramPart(part) || !partPaths.Add(PartPath(part))) return false;
            if (attribute.LocalName == "dm")
            {
                dataPart = (DiagramDataPart)part;
                dataRelationshipId = attribute.Value;
            }
        }
        if (dataPart is null || relationshipIds.Count != 4 || partPaths.Count != 4 ||
            expected.Count != rootAttributes.Select(attribute => attribute.LocalName).Distinct(StringComparer.Ordinal).Count())
            return false;

        byte[] sourceBytes;
        XDocument document;
        IReadOnlyList<DiagramNode> nodes;
        try
        {
            sourceBytes = ReadPart(dataPart);
            using var memory = new MemoryStream(sourceBytes, writable: false);
            using var reader = XmlReader.Create(memory, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreComments = false,
                IgnoreProcessingInstructions = false,
                IgnoreWhitespace = false,
            });
            document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
            if (!TryReadNodes(document, out nodes)) return false;
        }
        catch (Exception exception) when (exception is XmlException or IOException or UnauthorizedAccessException)
        {
            return false;
        }

        var binding = new PresentationDiagramText
        {
            PartPath = PartPath(dataPart),
            ContentType = dataPart.ContentType,
            SourceSha256 = Hash(sourceBytes),
            RelationshipId = dataRelationshipId,
        };
        binding.Nodes.Add(nodes.Select(node => new PresentationDiagramTextNode { ModelId = node.ModelId, Text = node.Text }));
        resolved = new ResolvedDiagram(binding, dataPart, document, nodes);
        return true;
    }

    private static bool TryReadNodes(XDocument document, out IReadOnlyList<DiagramNode> nodes)
    {
        nodes = [];
        var root = document.Root;
        if (root is null || root.Name.LocalName != "dataModel" || !DiagramNamespaces.Contains(root.Name.NamespaceName)) return false;
        var pointLists = root.Elements().Where(element => IsDiagram(element, "ptLst")).ToArray();
        if (pointLists.Length != 1) return false;
        var results = new List<DiagramNode>();
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var point in pointLists[0].Elements().Where(element => IsDiagram(element, "pt")))
        {
            if (point.Attribute("type")?.Value != "doc") continue;
            var modelId = point.Attribute("modelId")?.Value ?? string.Empty;
            if (!IsBoundedModelId(modelId) || !ids.Add(modelId)) return false;
            var textBodies = point.Elements().Where(element => IsDiagram(element, "t")).ToArray();
            if (textBodies.Length != 1 || !TryReadPlainText(textBodies[0], out var text, out var textElement)) return false;
            results.Add(new DiagramNode(modelId, text, textElement));
        }
        if (results.Count == 0) return false;
        nodes = results;
        return true;
    }

    private static bool TryReadPlainText(XElement body, out string text, out XElement textElement)
    {
        text = string.Empty;
        textElement = null!;
        var bodyChildren = body.Elements().ToArray();
        if (bodyChildren.Any(element => !IsDrawing(element, "bodyPr") && !IsDrawing(element, "lstStyle") && !IsDrawing(element, "p"))) return false;
        var paragraphs = bodyChildren.Where(element => IsDrawing(element, "p")).ToArray();
        if (paragraphs.Length != 1) return false;
        var paragraphChildren = paragraphs[0].Elements().ToArray();
        if (paragraphChildren.Any(element => !IsDrawing(element, "pPr") && !IsDrawing(element, "r") && !IsDrawing(element, "endParaRPr"))) return false;
        var runs = paragraphChildren.Where(element => IsDrawing(element, "r")).ToArray();
        if (runs.Length != 1) return false;
        var runChildren = runs[0].Elements().ToArray();
        if (runChildren.Any(element => !IsDrawing(element, "rPr") && !IsDrawing(element, "t"))) return false;
        var textElements = runChildren.Where(element => IsDrawing(element, "t")).ToArray();
        // Replacing XElement.Value would discard comments or processing
        // instructions nested in a:t. They are not part of the plain-text
        // profile, so withhold the capability rather than silently erasing
        // source-owned markup.
        if (textElements.Length != 1 || textElements[0].HasElements || textElements[0].Nodes().Any(node => node is not XText)) return false;
        text = textElements[0].Value;
        if (!IsBoundedText(text)) return false;
        textElement = textElements[0];
        return true;
    }

    private static void ValidateRequestedNodes(
        Google.Protobuf.Collections.RepeatedField<PresentationDiagramTextNode> original,
        Google.Protobuf.Collections.RepeatedField<PresentationDiagramTextNode> requested,
        string partPath)
    {
        if (original.Count != requested.Count)
            throw Unsupported("SmartArt node topology is source-bound and cannot be changed.", partPath);
        for (var index = 0; index < original.Count; index++)
        {
            if (original[index].ModelId != requested[index].ModelId)
                throw Unsupported("SmartArt node identifiers are source-bound and cannot be changed.", partPath);
            if (!IsBoundedText(requested[index].Text))
                throw Unsupported($"SmartArt node {requested[index].ModelId} text is outside the bounded plain-text profile.", partPath);
        }
    }

    private static bool SameBinding(PresentationDiagramText expected, PresentationDiagramText actual) =>
        expected.PartPath.Equals(actual.PartPath, StringComparison.OrdinalIgnoreCase) &&
        expected.ContentType.Equals(actual.ContentType, StringComparison.OrdinalIgnoreCase) &&
        expected.SourceSha256.Equals(actual.SourceSha256, StringComparison.OrdinalIgnoreCase) &&
        expected.RelationshipId == actual.RelationshipId;

    private static bool SameNodes(
        Google.Protobuf.Collections.RepeatedField<PresentationDiagramTextNode> expected,
        IReadOnlyList<DiagramNode> actual) =>
        expected.Count == actual.Count && expected.Select((node, index) =>
            node.ModelId == actual[index].ModelId && node.Text == actual[index].Text).All(match => match);

    private static bool SameNodeIds(
        Google.Protobuf.Collections.RepeatedField<PresentationDiagramTextNode> expected,
        IReadOnlyList<DiagramNode> actual) =>
        expected.Count == actual.Count && expected.Select((node, index) => node.ModelId == actual[index].ModelId).All(match => match);

    private static bool SameRequestedText(
        Google.Protobuf.Collections.RepeatedField<PresentationDiagramTextNode> expected,
        IReadOnlyList<DiagramNode> actual) =>
        expected.Count == actual.Count && expected.Select((node, index) =>
            node.ModelId == actual[index].ModelId && node.Text == actual[index].Text).All(match => match);

    private static bool IsClosedDiagramPart(OpenXmlPart part) =>
        part.Parts.Any() == false && !part.ExternalRelationships.Any() && !part.HyperlinkRelationships.Any() && !part.DataPartReferenceRelationships.Any();

    private static bool IsDiagram(XElement element, string localName) =>
        element.Name.LocalName == localName && DiagramNamespaces.Contains(element.Name.NamespaceName);

    private static bool IsDrawing(XElement element, string localName) =>
        element.Name.LocalName == localName && DrawingNamespaces.Contains(element.Name.NamespaceName);

    private static bool IsBoundedModelId(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > MaxModelIdLength || value.Any(char.IsControl)) return false;
        try
        {
            XmlConvert.VerifyXmlChars(value);
            // ST_ModelId is the ECMA-376 union of xsd:int and a:ST_Guid.
            // Accepting arbitrary XML-safe strings here would expose a write
            // capability for source packages that the Open XML validator must
            // later reject. Keep the import capability at the same boundary.
            return int.TryParse(value, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out _) ||
                Guid.TryParseExact(value, "B", out _);
        }
        catch (XmlException)
        {
            return false;
        }
    }

    private static bool IsBoundedText(string value)
    {
        if (value.Length > MaxNodeTextLength || value.Any(character => char.IsControl(character) && character is not '\t' and not '\n' and not '\r')) return false;
        try
        {
            XmlConvert.VerifyXmlChars(value);
            return true;
        }
        catch (XmlException)
        {
            return false;
        }
    }

    private static void SetText(XElement element, string value)
    {
        element.Value = value;
        var preserveWhitespace = value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1]));
        element.SetAttributeValue(XNamespace.Xml + "space", preserveWhitespace ? "preserve" : null);
    }

    private static byte[] Serialize(XDocument document)
    {
        using var stream = new MemoryStream();
        using (var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            NewLineHandling = NewLineHandling.None,
            OmitXmlDeclaration = document.Declaration is null,
        }))
        {
            document.Save(writer);
        }
        return stream.ToArray();
    }

    private static byte[] ReadPart(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }

    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static string Hash(byte[] data) => Convert.ToHexString(SHA256.HashData(data)).ToLowerInvariant();
    private static CodecException Unsupported(string message, string? partPath = null) => new("unsupported_presentation_edit", message, partPath);
    private static CodecException BindingMismatch(string message, string? partPath = null, Exception? innerException = null) =>
        new("presentation_diagram_text_binding_mismatch", message, partPath, innerException);
}
