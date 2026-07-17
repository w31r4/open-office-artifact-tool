using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns the bounded stylesheet projections. Existing resources, XF records,
// and differential formats are immutable: edits append derived records and
// preserve every pre-existing record plus unmodeled stylesheet content.
internal sealed class XlsxCellStyleCodec
{
    private static readonly HashSet<string> PatternTypes = new(StringComparer.Ordinal)
    {
        "none", "solid", "mediumGray", "darkGray", "lightGray", "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis",
        "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid", "lightTrellis", "gray125", "gray0625",
    };

    private static readonly HashSet<string> BorderStyles = new(StringComparer.Ordinal)
    {
        "none", "thin", "medium", "dashed", "dotted", "thick", "double", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot",
    };

    private static readonly HashSet<string> UnderlineStyles = new(StringComparer.Ordinal)
    {
        "single", "double", "singleAccounting", "doubleAccounting", "none",
    };

    private static readonly HashSet<string> HorizontalAlignments = new(StringComparer.Ordinal)
    {
        "general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed",
    };

    private static readonly HashSet<string> VerticalAlignments = new(StringComparer.Ordinal)
    {
        "top", "center", "bottom", "justify", "distributed",
    };

    private readonly WorkbookPart _workbookPart;
    private WorkbookStylesPart? _stylesPart;
    private Stylesheet? _stylesheet;
    private readonly Dictionary<uint, string> _customFormatsById = [];
    private readonly Dictionary<string, uint> _customFormatIds = new(StringComparer.Ordinal);
    private List<Font> _fonts = [];
    private List<Fill> _fills = [];
    private List<Border> _borders = [];
    private List<CellFormat> _cellFormats = [];
    private List<DifferentialFormat> _differentialFormats = [];
    private Dictionary<string, uint> _fontIds = new(StringComparer.Ordinal);
    private Dictionary<string, uint> _fillIds = new(StringComparer.Ordinal);
    private Dictionary<string, uint> _borderIds = new(StringComparer.Ordinal);
    private Dictionary<string, uint> _cellFormatIds = new(StringComparer.Ordinal);
    private Dictionary<string, uint> _differentialFormatIds = new(StringComparer.Ordinal);
    private string[] _originalNumberingFormats = [];
    private string[] _originalFonts = [];
    private string[] _originalFills = [];
    private string[] _originalBorders = [];
    private string[] _originalCellFormats = [];
    private string[] _originalDifferentialFormats = [];
    private string[] _originalOtherChildren = [];
    private bool _dirty;

    internal XlsxCellStyleCodec(WorkbookPart workbookPart)
    {
        _workbookPart = workbookPart;
        _stylesPart = workbookPart.WorkbookStylesPart;
        _stylesheet = _stylesPart?.Stylesheet;
        if (_stylesPart is not null && _stylesheet is null)
            throw Invalid("Workbook styles part has no stylesheet root.");
        if (_stylesheet is not null) IndexExistingStyles(_stylesheet);
    }

    internal string ReadNumberFormat(Cell cell)
    {
        var format = CellFormatFor(cell);
        if (format is null) return string.Empty;
        var numberFormatId = EffectiveNumberFormatId(format);
        if (XlsxNumberFormatCodec.TryGetBuiltInFormat(numberFormatId, out var builtIn)) return builtIn;
        if (_customFormatsById.TryGetValue(numberFormatId, out var custom)) return custom;
        if (numberFormatId < 164)
            throw XlsxNumberFormatCodec.Invalid($"Workbook uses locale-dependent built-in number format {numberFormatId}, which this codec cannot represent as a stable format code.");
        throw XlsxNumberFormatCodec.Invalid($"Workbook references missing custom number format {numberFormatId}.");
    }

    internal CellStyleArtifact? ReadStyle(Cell cell)
    {
        var format = CellFormatFor(cell);
        if (format is null || _cellFormats.Count == 0) return null;
        var current = ReadFullStyle(format);
        var baseline = ReadFullStyle(_cellFormats[0]);
        var result = new CellStyleArtifact();
        if (!Equals(current.Font, baseline.Font)) result.Font = current.Font;
        if (!Equals(current.Fill, baseline.Fill)) result.Fill = current.Fill;
        if (!Equals(current.Border, baseline.Border)) result.Border = current.Border;
        if (!Equals(current.Alignment, baseline.Alignment)) result.Alignment = current.Alignment;
        if (!Equals(current.Protection, baseline.Protection)) result.Protection = current.Protection;
        return HasStyle(result) ? result : null;
    }

    internal bool TryReadTableColor(uint differentialFormatId, bool cellColor, out SpreadsheetTableColorArtifact? artifact)
    {
        artifact = null;
        if (differentialFormatId >= _differentialFormats.Count) return false;
        try
        {
            var differential = _differentialFormats[checked((int)differentialFormatId)];
            if (differential.ChildElements.Count != 1) return false;
            SpreadsheetColor? color;
            if (cellColor)
            {
                if (differential.ChildElements[0] is not Fill fill || fill.ChildElements.Count != 1 || fill.PatternFill is not { } pattern ||
                    pattern.PatternType?.Value != PatternValues.Solid || pattern.ForegroundColor is null) return false;
                var children = pattern.ChildElements.ToArray();
                if (children.Length is < 1 or > 2 || children[0] is not ForegroundColor) return false;
                if (children.Length == 2 && (children[1] is not BackgroundColor background || !IsAutomaticPatternBackground(background))) return false;
                color = ReadColor(pattern.ForegroundColor);
            }
            else
            {
                if (differential.ChildElements[0] is not Font font || font.ChildElements.Count != 1 ||
                    font.ChildElements[0] is not DocumentFormat.OpenXml.Spreadsheet.Color fontColor) return false;
                color = ReadColor(fontColor);
            }
            if (color is null) return false;
            artifact = new SpreadsheetTableColorArtifact { Color = color };
            if (cellColor) artifact.CellColor = true;
            else artifact.FontColor = true;
            return true;
        }
        catch (CodecException)
        {
            return false;
        }
    }

    internal bool TryReadDifferentialStyle(uint differentialFormatId, out CellStyleArtifact? artifact)
    {
        artifact = null;
        if (differentialFormatId >= _differentialFormats.Count) return false;
        try
        {
            var differential = _differentialFormats[checked((int)differentialFormatId)];
            if (!HasOnlyAttributes(differential) || differential.ChildElements.Any(item => item is not Font and not Fill and not Alignment and not Border and not Protection) ||
                differential.Elements<Font>().Skip(1).Any() || differential.Elements<Fill>().Skip(1).Any() ||
                differential.Elements<Alignment>().Skip(1).Any() || differential.Elements<Border>().Skip(1).Any() ||
                differential.Elements<Protection>().Skip(1).Any()) return false;

            var result = new CellStyleArtifact();
            if (differential.GetFirstChild<Font>() is { } font)
            {
                if (!TryReadDifferentialFont(font, out var value) || value is null) return false;
                result.Font = value;
            }
            if (differential.GetFirstChild<Fill>() is { } fill)
            {
                if (!IsBoundedDifferentialFill(fill) || ReadFill(fill) is not { } value) return false;
                result.Fill = value;
            }
            if (differential.GetFirstChild<Alignment>() is { } alignment)
            {
                if (!IsBoundedDifferentialAlignment(alignment) || ReadAlignment(alignment) is not { } value) return false;
                result.Alignment = value;
            }
            if (differential.GetFirstChild<Border>() is { } border)
            {
                if (!IsBoundedDifferentialBorder(border) || ReadBorder(border) is not { } value) return false;
                result.Border = value;
            }
            if (differential.GetFirstChild<Protection>() is { } protection)
            {
                if (!IsBoundedDifferentialProtection(protection) || ReadProtection(protection) is not { } value) return false;
                result.Protection = value;
            }
            Validate(result, $"dxf {differentialFormatId}");
            if (!HasStyle(result)) return false;
            artifact = result;
            return true;
        }
        catch (CodecException)
        {
            artifact = null;
            return false;
        }
    }

    internal uint FindOrCreateDifferentialStyle(CellStyleArtifact source, string location)
    {
        Validate(source, location);
        EnsureWritableStylesheet();
        var differential = new DifferentialFormat();
        if (source.Font is not null) differential.Append(ApplyFont(new Font(), source.Font));
        if (source.Fill is not null) differential.Append(ApplyDifferentialFill(new Fill(), source.Fill));
        if (source.Alignment is not null) differential.Append(ApplyAlignment(new Alignment(), source.Alignment));
        if (source.Border is not null) differential.Append(ApplyBorder(new Border(), source.Border));
        if (source.Protection is not null) differential.Append(ApplyProtection(new Protection(), source.Protection));
        var key = differential.OuterXml;
        if (_differentialFormatIds.TryGetValue(key, out var existing)) return existing;
        var collection = _stylesheet!.DifferentialFormats;
        if (collection is null)
        {
            collection = new DifferentialFormats();
            var before = _stylesheet.ChildElements.FirstOrDefault(item => item.LocalName is "tableStyles" or "colors" or "extLst");
            if (before is null) _stylesheet.Append(collection);
            else _stylesheet.InsertBefore(collection, before);
        }
        collection.Append(differential);
        _differentialFormats.Add(differential);
        collection.Count = checked((uint)_differentialFormats.Count);
        var id = checked((uint)_differentialFormats.Count - 1);
        _differentialFormatIds[key] = id;
        _dirty = true;
        return id;
    }

    internal static SpreadsheetColor? ReadConditionalColor(DocumentFormat.OpenXml.Spreadsheet.Color color) => ReadColor(color);

    internal static DocumentFormat.OpenXml.Spreadsheet.Color WriteConditionalColor(SpreadsheetColor source, string location)
    {
        ValidateColor(source, location, "conditional format");
        return ApplyColor(new DocumentFormat.OpenXml.Spreadsheet.Color(), source);
    }

    internal uint FindOrCreateTableColor(SpreadsheetTableColorArtifact source, string location)
    {
        ValidateTableColor(source, location);
        EnsureWritableStylesheet();
        var differential = new DifferentialFormat();
        if (source.TargetCase == SpreadsheetTableColorArtifact.TargetOneofCase.CellColor)
        {
            var pattern = new PatternFill { PatternType = PatternValues.Solid };
            pattern.Append(ApplyColor(new ForegroundColor(), source.Color));
            pattern.Append(new BackgroundColor { Indexed = 64U });
            differential.Append(new Fill(pattern));
        }
        else
        {
            differential.Append(new Font(ApplyColor(new DocumentFormat.OpenXml.Spreadsheet.Color(), source.Color)));
        }
        var key = differential.OuterXml;
        if (_differentialFormatIds.TryGetValue(key, out var existing)) return existing;
        var collection = _stylesheet!.DifferentialFormats;
        if (collection is null)
        {
            collection = new DifferentialFormats();
            var before = _stylesheet.ChildElements.FirstOrDefault(item => item.LocalName is "tableStyles" or "colors" or "extLst");
            if (before is null) _stylesheet.Append(collection);
            else _stylesheet.InsertBefore(collection, before);
        }
        collection.Append(differential);
        _differentialFormats.Add(differential);
        collection.Count = checked((uint)_differentialFormats.Count);
        var id = checked((uint)_differentialFormats.Count - 1);
        _differentialFormatIds[key] = id;
        _dirty = true;
        return id;
    }

    internal static void ValidateTableColor(SpreadsheetTableColorArtifact? source, string location)
    {
        if (source is null || source.Color is null ||
            source.TargetCase == SpreadsheetTableColorArtifact.TargetOneofCase.None ||
            source.TargetCase == SpreadsheetTableColorArtifact.TargetOneofCase.CellColor && !source.CellColor ||
            source.TargetCase == SpreadsheetTableColorArtifact.TargetOneofCase.FontColor && !source.FontColor)
            throw InvalidTableColor($"Worksheet table {location} color selector must provide exactly one true target and one color.");
        try
        {
            ValidateColor(source.Color, location, "table");
        }
        catch (CodecException)
        {
            throw InvalidTableColor($"Worksheet table {location} has an invalid color selector.");
        }
    }

    internal void Apply(Cell cell, CellArtifact source)
    {
        var location = cell.CellReference?.Value ?? $"R{source.Row + 1}C{source.Column + 1}";
        var desiredNumberFormat = XlsxNumberFormatCodec.Canonicalize(source.NumberFormatCode, location);
        Validate(source.Style, location);
        var currentNumberFormat = ReadNumberFormat(cell);
        var currentStyle = ReadStyle(cell);
        if (string.Equals(currentNumberFormat, desiredNumberFormat, StringComparison.Ordinal) && Equals(currentStyle, source.Style)) return;

        EnsureWritableStylesheet();
        var sourceIndex = cell.StyleIndex?.Value ?? 0;
        if (sourceIndex >= _cellFormats.Count) throw Invalid($"Cell {location} references missing style {sourceIndex}.");
        var derived = (CellFormat)_cellFormats[checked((int)sourceIndex)].CloneNode(true);

        if (!string.Equals(currentNumberFormat, desiredNumberFormat, StringComparison.Ordinal))
        {
            derived.NumberFormatId = FindOrCreateFormatId(desiredNumberFormat);
            derived.ApplyNumberFormat = true;
        }

        var desired = source.Style;
        if (!Equals(currentStyle?.Font, desired?.Font))
        {
            derived.FontId = desired?.Font is null ? DefaultFontId() : FindOrCreateFont(ApplyFont(CloneCurrentFont(derived), desired.Font));
            derived.ApplyFont = true;
        }
        if (!Equals(currentStyle?.Fill, desired?.Fill))
        {
            derived.FillId = desired?.Fill is null ? DefaultFillId() : FindOrCreateFill(ApplyFill(CloneCurrentFill(derived), desired.Fill));
            derived.ApplyFill = true;
        }
        if (!Equals(currentStyle?.Border, desired?.Border))
        {
            derived.BorderId = desired?.Border is null ? DefaultBorderId() : FindOrCreateBorder(ApplyBorder(CloneCurrentBorder(derived), desired.Border));
            derived.ApplyBorder = true;
        }
        if (!Equals(currentStyle?.Alignment, desired?.Alignment))
        {
            if (desired?.Alignment is null) derived.Alignment = null;
            else derived.Alignment = ApplyAlignment(derived.Alignment?.CloneNode(true) as Alignment ?? new Alignment(), desired.Alignment);
            derived.ApplyAlignment = true;
        }
        if (!Equals(currentStyle?.Protection, desired?.Protection))
        {
            if (desired?.Protection is null) derived.Protection = null;
            else derived.Protection = ApplyProtection(derived.Protection?.CloneNode(true) as Protection ?? new Protection(), desired.Protection);
            derived.ApplyProtection = true;
        }

        cell.StyleIndex = FindOrCreateCellFormat(derived);
        _dirty = true;
    }

    internal void Save()
    {
        if (!_dirty || _stylesheet is null) return;
        AssertOriginalStylesPreserved();
        _stylesheet.Save();
    }

    internal static void Validate(CellStyleArtifact? style, string location)
    {
        if (style is null) return;
        if (style.Font is { } font)
        {
            if (font.HasUnderline && !UnderlineStyles.Contains(font.Underline)) throw Invalid($"Cell {location} has unsupported underline style {font.Underline}.");
            if (font.HasSizePoints && (!double.IsFinite(font.SizePoints) || font.SizePoints <= 0 || font.SizePoints > 409)) throw Invalid($"Cell {location} font size must be greater than zero and at most 409 points.");
            if (font.HasName && (string.IsNullOrWhiteSpace(font.Name) || font.Name.Length > 255 || font.Name.Any(char.IsControl))) throw Invalid($"Cell {location} has an invalid font name.");
            ValidateColor(font.Color, location, "font");
        }
        if (style.Fill is { } fill)
        {
            if (!PatternTypes.Contains(fill.PatternType)) throw Invalid($"Cell {location} has unsupported fill pattern {fill.PatternType}.");
            if (fill.PatternType == "solid" && fill.Foreground is null) throw Invalid($"Cell {location} solid fill requires a foreground color.");
            ValidateColor(fill.Foreground, location, "fill foreground");
            ValidateColor(fill.Background, location, "fill background");
        }
        if (style.Border is { } border)
        {
            foreach (var (name, edge) in BorderEdges(border))
            {
                if (edge is null) continue;
                if (!BorderStyles.Contains(edge.Style) || edge.Style == "none") throw Invalid($"Cell {location} has unsupported {name} border style {edge.Style}.");
                ValidateColor(edge.Color, location, $"{name} border");
            }
        }
        if (style.Alignment is { } alignment)
        {
            if (alignment.HasHorizontal && !HorizontalAlignments.Contains(alignment.Horizontal)) throw Invalid($"Cell {location} has unsupported horizontal alignment {alignment.Horizontal}.");
            if (alignment.HasVertical && !VerticalAlignments.Contains(alignment.Vertical)) throw Invalid($"Cell {location} has unsupported vertical alignment {alignment.Vertical}.");
            if (alignment.HasTextRotation && alignment.TextRotation > 180 && alignment.TextRotation != 255) throw Invalid($"Cell {location} text rotation must be 0 through 180 or 255.");
            if (alignment.HasIndent && alignment.Indent > 250) throw Invalid($"Cell {location} indent must be at most 250.");
            if (alignment.HasReadingOrder && alignment.ReadingOrder > 2) throw Invalid($"Cell {location} reading order must be 0, 1, or 2.");
        }
    }

    private CellFormat? CellFormatFor(Cell cell)
    {
        if (cell.StyleIndex?.HasValue != true) return _cellFormats.Count == 0 ? null : _cellFormats[0];
        if (_stylesheet?.CellFormats is null) throw XlsxNumberFormatCodec.Invalid($"Cell {cell.CellReference} references style {cell.StyleIndex.Value}, but the workbook has no cell formats.");
        var styleIndex = cell.StyleIndex.Value;
        if (styleIndex >= _cellFormats.Count) throw XlsxNumberFormatCodec.Invalid($"Cell {cell.CellReference} references missing style {styleIndex}.");
        return _cellFormats[checked((int)styleIndex)];
    }

    private CellStyleArtifact ReadFullStyle(CellFormat format)
    {
        var target = new CellStyleArtifact();
        var fontId = EffectiveResourceId(format, item => item.FontId?.Value ?? 0, format.ApplyFont);
        var fillId = EffectiveResourceId(format, item => item.FillId?.Value ?? 0, format.ApplyFill);
        var borderId = EffectiveResourceId(format, item => item.BorderId?.Value ?? 0, format.ApplyBorder);
        if (fontId >= _fonts.Count) throw Invalid($"Cell format references missing font {fontId}.");
        if (fillId >= _fills.Count) throw Invalid($"Cell format references missing fill {fillId}.");
        if (borderId >= _borders.Count) throw Invalid($"Cell format references missing border {borderId}.");
        target.Font = ReadFont(_fonts[checked((int)fontId)]);
        target.Fill = ReadFill(_fills[checked((int)fillId)]);
        target.Border = ReadBorder(_borders[checked((int)borderId)]);
        var baseFormat = BaseFormat(format);
        var alignment = format.ApplyAlignment?.Value == false ? baseFormat?.Alignment : format.Alignment ?? baseFormat?.Alignment;
        var protection = format.ApplyProtection?.Value == false ? baseFormat?.Protection : format.Protection ?? baseFormat?.Protection;
        target.Alignment = ReadAlignment(alignment);
        target.Protection = ReadProtection(protection);
        return target;
    }

    private uint EffectiveNumberFormatId(CellFormat format) => EffectiveResourceId(format, item => item.NumberFormatId?.Value ?? 0, format.ApplyNumberFormat);

    private uint EffectiveResourceId(CellFormat format, Func<CellFormat, uint> selector, BooleanValue? apply)
    {
        var direct = selector(format);
        var baseFormat = BaseFormat(format);
        if (baseFormat is null) return direct;
        if (apply?.Value == false) return selector(baseFormat);
        var inherited = selector(baseFormat);
        return apply?.Value == true || direct != inherited ? direct : inherited;
    }

    private CellFormat? BaseFormat(CellFormat format)
    {
        var formatId = format.FormatId?.Value;
        if (formatId is null || _stylesheet?.CellStyleFormats is null) return null;
        var bases = _stylesheet.CellStyleFormats.Elements<CellFormat>().ToArray();
        if (formatId.Value >= bases.Length) throw Invalid($"Cell format references missing base style {formatId.Value}.");
        return bases[checked((int)formatId.Value)];
    }

    private static SpreadsheetFontStyle ReadFont(Font font)
    {
        var target = new SpreadsheetFontStyle
        {
            Bold = ElementBoolean(font.GetFirstChild<Bold>()),
            Italic = ElementBoolean(font.GetFirstChild<Italic>()),
            Strike = ElementBoolean(font.GetFirstChild<Strike>()),
            SizePoints = font.GetFirstChild<FontSize>()?.Val?.Value ?? 11,
            Name = font.GetFirstChild<FontName>()?.Val?.Value ?? "Aptos",
        };
        if (font.GetFirstChild<Underline>() is { } underline)
        {
            var underlineText = underline.Val is { } underlineValue ? (string?)underlineValue : null;
            target.Underline = underlineText ?? "single";
        }
        target.Color = ReadColor(font.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>());
        return target;
    }

    private static bool TryReadDifferentialFont(Font font, out SpreadsheetFontStyle? artifact)
    {
        artifact = null;
        if (!HasOnlyAttributes(font) || font.ChildElements.Any(item => item is not Bold and not Italic and not Strike and not Underline and not FontSize and not FontName and not DocumentFormat.OpenXml.Spreadsheet.Color) ||
            font.Elements<Bold>().Skip(1).Any() || font.Elements<Italic>().Skip(1).Any() || font.Elements<Strike>().Skip(1).Any() ||
            font.Elements<Underline>().Skip(1).Any() || font.Elements<FontSize>().Skip(1).Any() || font.Elements<FontName>().Skip(1).Any() ||
            font.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().Skip(1).Any()) return false;
        if (font.ChildElements.Any(item => item switch
            {
                Bold or Italic or Strike or Underline or FontSize or FontName => item.HasChildren || !HasOnlyAttributes(item, "val"),
                DocumentFormat.OpenXml.Spreadsheet.Color color => !IsBoundedDifferentialColor(color),
                _ => true,
            })) return false;
        var target = new SpreadsheetFontStyle();
        if (font.GetFirstChild<Bold>() is { } bold) target.Bold = ElementBoolean(bold);
        if (font.GetFirstChild<Italic>() is { } italic) target.Italic = ElementBoolean(italic);
        if (font.GetFirstChild<Strike>() is { } strike) target.Strike = ElementBoolean(strike);
        if (font.GetFirstChild<Underline>() is { } underline)
        {
            var underlineText = underline.Val is { } underlineValue ? (string?)underlineValue : null;
            target.Underline = underlineText ?? "single";
        }
        if (font.GetFirstChild<FontSize>()?.Val?.HasValue == true) target.SizePoints = font.FontSize!.Val!.Value;
        if (font.GetFirstChild<FontName>()?.Val?.HasValue == true) target.Name = font.FontName!.Val!.Value!;
        target.Color = ReadColor(font.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>());
        artifact = target.HasBold || target.HasItalic || target.HasStrike || target.HasUnderline || target.HasSizePoints || target.HasName || target.Color is not null ? target : null;
        return true;
    }

    private static bool IsBoundedDifferentialFill(Fill fill)
    {
        if (!HasOnlyAttributes(fill) || fill.ChildElements.Count != 1 || fill.PatternFill is not { } pattern ||
            !HasOnlyAttributes(pattern, "patternType") || pattern.ChildElements.Any(item => item is not ForegroundColor and not BackgroundColor) ||
            pattern.Elements<ForegroundColor>().Skip(1).Any() || pattern.Elements<BackgroundColor>().Skip(1).Any()) return false;
        return pattern.ChildElements.All(item => item is ColorType color && IsBoundedDifferentialColor(color));
    }

    private static bool IsBoundedDifferentialAlignment(Alignment alignment) =>
        !alignment.HasChildren && HasOnlyAttributes(alignment, "horizontal", "vertical", "wrapText", "textRotation", "indent", "shrinkToFit", "readingOrder");

    private static bool IsBoundedDifferentialBorder(Border border)
    {
        if (!HasOnlyAttributes(border, "diagonalUp", "diagonalDown", "outline")) return false;
        var allowed = new HashSet<string>(StringComparer.Ordinal) { "left", "right", "top", "bottom", "diagonal", "start", "end", "horizontal", "vertical" };
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var edge in border.ChildElements)
        {
            if (edge is not BorderPropertiesType || !allowed.Contains(edge.LocalName) || !names.Add(edge.LocalName) || !HasOnlyAttributes(edge, "style") ||
                edge.ChildElements.Any(item => item is not DocumentFormat.OpenXml.Spreadsheet.Color) ||
                edge.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().Skip(1).Any() ||
                edge.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>() is { } color && !IsBoundedDifferentialColor(color)) return false;
        }
        return true;
    }

    private static bool IsBoundedDifferentialProtection(Protection protection) =>
        !protection.HasChildren && HasOnlyAttributes(protection, "locked", "hidden");

    private static bool IsBoundedDifferentialColor(ColorType color) =>
        !color.HasChildren && HasOnlyAttributes(color, "rgb", "theme", "indexed", "auto", "tint");

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = new HashSet<string>(names, StringComparer.Ordinal);
        return element.GetAttributes().All(attribute => allowed.Contains(attribute.LocalName));
    }

    private static SpreadsheetFillStyle? ReadFill(Fill fill)
    {
        var pattern = fill.PatternFill;
        if (pattern is null) return null;
        var patternText = pattern.PatternType is { } patternValue ? (string?)patternValue : null;
        var patternType = patternText ?? "none";
        if (patternType == "none") return null;
        var foreground = ReadColor(pattern.ForegroundColor);
        var background = ReadColor(pattern.BackgroundColor);
        if (patternType == "solid" && (background?.SourceCase == SpreadsheetColor.SourceOneofCase.Indexed && background.Indexed == 64 ||
            foreground is not null && background is not null && foreground.Equals(background))) background = null;
        return new SpreadsheetFillStyle
        {
            PatternType = patternType,
            Foreground = foreground,
            Background = background,
        };
    }

    private static SpreadsheetBorderStyle? ReadBorder(Border border)
    {
        var target = new SpreadsheetBorderStyle
        {
            Left = ReadBorderEdge(border.LeftBorder), Right = ReadBorderEdge(border.RightBorder),
            Top = ReadBorderEdge(border.TopBorder), Bottom = ReadBorderEdge(border.BottomBorder),
            Diagonal = ReadBorderEdge(border.DiagonalBorder), Start = ReadBorderEdge(border.StartBorder), End = ReadBorderEdge(border.EndBorder),
            Horizontal = ReadBorderEdge(border.HorizontalBorder), Vertical = ReadBorderEdge(border.VerticalBorder),
        };
        if (border.DiagonalUp?.HasValue == true) target.DiagonalUp = border.DiagonalUp.Value;
        if (border.DiagonalDown?.HasValue == true) target.DiagonalDown = border.DiagonalDown.Value;
        if (border.Outline?.HasValue == true) target.Outline = border.Outline.Value;
        return HasBorder(target) ? target : null;
    }

    private static SpreadsheetBorderEdgeStyle? ReadBorderEdge(BorderPropertiesType? edge)
    {
        var style = edge?.Style is { } styleValue ? (string?)styleValue : null;
        if (string.IsNullOrEmpty(style) || style == "none") return null;
        return new SpreadsheetBorderEdgeStyle { Style = style, Color = ReadColor(edge!.Color) };
    }

    private static SpreadsheetAlignmentStyle? ReadAlignment(Alignment? alignment)
    {
        if (alignment is null) return null;
        var target = new SpreadsheetAlignmentStyle();
        if (alignment.Horizontal?.HasValue == true) target.Horizontal = (string?)alignment.Horizontal ?? string.Empty;
        if (alignment.Vertical?.HasValue == true) target.Vertical = (string?)alignment.Vertical ?? string.Empty;
        if (alignment.WrapText?.HasValue == true) target.WrapText = alignment.WrapText.Value;
        if (alignment.TextRotation?.HasValue == true) target.TextRotation = alignment.TextRotation.Value;
        if (alignment.Indent?.HasValue == true) target.Indent = alignment.Indent.Value;
        if (alignment.ShrinkToFit?.HasValue == true) target.ShrinkToFit = alignment.ShrinkToFit.Value;
        if (alignment.ReadingOrder?.HasValue == true) target.ReadingOrder = alignment.ReadingOrder.Value;
        return HasAlignment(target) ? target : null;
    }

    private static SpreadsheetProtectionStyle? ReadProtection(Protection? protection)
    {
        if (protection is null) return null;
        var target = new SpreadsheetProtectionStyle();
        if (protection.Locked?.HasValue == true) target.Locked = protection.Locked.Value;
        if (protection.Hidden?.HasValue == true) target.Hidden = protection.Hidden.Value;
        return target.HasLocked || target.HasHidden ? target : null;
    }

    private static SpreadsheetColor? ReadColor(ColorType? color)
    {
        if (color is null) return null;
        var sources = (color.Rgb?.HasValue == true ? 1 : 0) + (color.Theme?.HasValue == true ? 1 : 0) + (color.Indexed?.HasValue == true ? 1 : 0) + (color.Auto?.Value == true ? 1 : 0);
        if (sources == 0) return null;
        if (sources != 1) throw Invalid("Spreadsheet color must use exactly one of rgb, theme, indexed, or auto.");
        var target = new SpreadsheetColor();
        if (color.Rgb?.HasValue == true)
        {
            var rgb = color.Rgb.Value ?? string.Empty;
            if (rgb.Length is not 6 and not 8 || !rgb.All(Uri.IsHexDigit)) throw Invalid($"Spreadsheet color has invalid RGB value {rgb}.");
            target.Rgb = rgb[^6..].ToUpperInvariant();
        }
        else if (color.Theme?.HasValue == true) target.Theme = color.Theme.Value;
        else if (color.Indexed?.HasValue == true) target.Indexed = color.Indexed.Value;
        else target.Automatic = true;
        if (color.Tint?.HasValue == true && color.Tint.Value != 0) target.Tint = color.Tint.Value;
        ValidateColor(target, "styles.xml", "source");
        return target;
    }

    private Font CloneCurrentFont(CellFormat format)
    {
        var id = EffectiveResourceId(format, item => item.FontId?.Value ?? 0, format.ApplyFont);
        if (id >= _fonts.Count) throw Invalid($"Cell format references missing font {id}.");
        return (Font)_fonts[checked((int)id)].CloneNode(true);
    }

    private Fill CloneCurrentFill(CellFormat format)
    {
        var id = EffectiveResourceId(format, item => item.FillId?.Value ?? 0, format.ApplyFill);
        if (id >= _fills.Count) throw Invalid($"Cell format references missing fill {id}.");
        return (Fill)_fills[checked((int)id)].CloneNode(true);
    }

    private Border CloneCurrentBorder(CellFormat format)
    {
        var id = EffectiveResourceId(format, item => item.BorderId?.Value ?? 0, format.ApplyBorder);
        if (id >= _borders.Count) throw Invalid($"Cell format references missing border {id}.");
        return (Border)_borders[checked((int)id)].CloneNode(true);
    }

    private static Font ApplyFont(Font target, SpreadsheetFontStyle source)
    {
        target.Bold = source.HasBold && source.Bold ? new Bold() : null;
        target.Italic = source.HasItalic && source.Italic ? new Italic() : null;
        target.Strike = source.HasStrike && source.Strike ? new Strike() : null;
        target.Underline = source.HasUnderline && source.Underline != "none" ? new Underline { Val = new UnderlineValues(source.Underline) } : null;
        target.FontSize = source.HasSizePoints ? new FontSize { Val = source.SizePoints } : null;
        target.Color = source.Color is not null ? ApplyColor(new DocumentFormat.OpenXml.Spreadsheet.Color(), source.Color) : null;
        target.FontName = source.HasName ? new FontName { Val = source.Name } : null;
        return target;
    }

    private static Fill ApplyFill(Fill target, SpreadsheetFillStyle source)
    {
        foreach (var child in target.ChildElements.Where(item => item is PatternFill or GradientFill).ToArray()) child.Remove();
        var pattern = new PatternFill { PatternType = new PatternValues(source.PatternType) };
        if (source.Foreground is not null) pattern.ForegroundColor = ApplyColor(new ForegroundColor(), source.Foreground);
        if (source.Background is not null) pattern.BackgroundColor = ApplyColor(new BackgroundColor(), source.Background);
        else if (source.PatternType == "solid") pattern.BackgroundColor = new BackgroundColor { Indexed = 64U };
        target.Append(pattern);
        return target;
    }

    // LibreOffice applies a differential solid fill's background color when
    // printing/rendering. Duplicate an otherwise implicit foreground color
    // only in DXFs, then normalize it again on import so the public model
    // retains the conventional single-color solid-fill representation.
    private static Fill ApplyDifferentialFill(Fill target, SpreadsheetFillStyle source)
    {
        var fill = ApplyFill(target, source);
        if (source.PatternType == "solid" && source.Foreground is not null && source.Background is null && fill.PatternFill is { } pattern)
            pattern.BackgroundColor = ApplyColor(new BackgroundColor(), source.Foreground);
        return fill;
    }

    private static Border ApplyBorder(Border target, SpreadsheetBorderStyle source)
    {
        target.LeftBorder = ApplyBorderEdge(new LeftBorder(), source.Left);
        target.RightBorder = ApplyBorderEdge(new RightBorder(), source.Right);
        target.TopBorder = ApplyBorderEdge(new TopBorder(), source.Top);
        target.BottomBorder = ApplyBorderEdge(new BottomBorder(), source.Bottom);
        target.DiagonalBorder = ApplyBorderEdge(new DiagonalBorder(), source.Diagonal);
        target.StartBorder = ApplyBorderEdge(new StartBorder(), source.Start);
        target.EndBorder = ApplyBorderEdge(new EndBorder(), source.End);
        target.HorizontalBorder = ApplyBorderEdge(new HorizontalBorder(), source.Horizontal);
        target.VerticalBorder = ApplyBorderEdge(new VerticalBorder(), source.Vertical);
        target.DiagonalUp = source.HasDiagonalUp ? source.DiagonalUp : null;
        target.DiagonalDown = source.HasDiagonalDown ? source.DiagonalDown : null;
        target.Outline = source.HasOutline ? source.Outline : null;
        return target;
    }

    private static T ApplyBorderEdge<T>(T target, SpreadsheetBorderEdgeStyle? source) where T : BorderPropertiesType
    {
        if (source is null) return target;
        target.Style = new BorderStyleValues(source.Style);
        if (source.Color is not null) target.Color = ApplyColor(new DocumentFormat.OpenXml.Spreadsheet.Color(), source.Color);
        return target;
    }

    private static Alignment ApplyAlignment(Alignment target, SpreadsheetAlignmentStyle source)
    {
        target.Horizontal = source.HasHorizontal ? new HorizontalAlignmentValues(source.Horizontal) : null;
        target.Vertical = source.HasVertical ? new VerticalAlignmentValues(source.Vertical) : null;
        target.WrapText = source.HasWrapText ? source.WrapText : null;
        target.TextRotation = source.HasTextRotation ? source.TextRotation : null;
        target.Indent = source.HasIndent ? source.Indent : null;
        target.ShrinkToFit = source.HasShrinkToFit ? source.ShrinkToFit : null;
        target.ReadingOrder = source.HasReadingOrder ? source.ReadingOrder : null;
        return target;
    }

    private static Protection ApplyProtection(Protection target, SpreadsheetProtectionStyle source)
    {
        target.Locked = source.HasLocked ? source.Locked : null;
        target.Hidden = source.HasHidden ? source.Hidden : null;
        return target;
    }

    private static T ApplyColor<T>(T target, SpreadsheetColor source) where T : ColorType
    {
        target.Rgb = null;
        target.Theme = null;
        target.Indexed = null;
        target.Auto = null;
        target.Tint = source.HasTint ? source.Tint : null;
        switch (source.SourceCase)
        {
            case SpreadsheetColor.SourceOneofCase.Rgb: target.Rgb = $"FF{source.Rgb.ToUpperInvariant()}"; break;
            case SpreadsheetColor.SourceOneofCase.Theme: target.Theme = source.Theme; break;
            case SpreadsheetColor.SourceOneofCase.Indexed: target.Indexed = source.Indexed; break;
            case SpreadsheetColor.SourceOneofCase.Automatic: target.Auto = true; break;
            default: throw Invalid("Spreadsheet color has no source.");
        }
        return target;
    }

    private uint FindOrCreateFormatId(string formatCode)
    {
        if (formatCode.Length == 0) return 0;
        if (XlsxNumberFormatCodec.TryGetBuiltInId(formatCode, out var builtInId)) return builtInId;
        if (_customFormatIds.TryGetValue(formatCode, out var existingId)) return existingId;
        var nextId = Math.Max(164U, _customFormatsById.Keys.DefaultIfEmpty(163U).Max() + 1U);
        var numberingFormats = _stylesheet!.NumberingFormats;
        if (numberingFormats is null)
        {
            numberingFormats = new NumberingFormats();
            if (_stylesheet.Fonts is { } fonts) _stylesheet.InsertBefore(numberingFormats, fonts);
            else _stylesheet.PrependChild(numberingFormats);
        }
        numberingFormats.Append(new NumberingFormat { NumberFormatId = nextId, FormatCode = formatCode });
        numberingFormats.Count = checked((uint)numberingFormats.ChildElements.Count);
        _customFormatsById[nextId] = formatCode;
        _customFormatIds[formatCode] = nextId;
        _dirty = true;
        return nextId;
    }

    private uint FindOrCreateFont(Font font) => FindOrAppend(font, _stylesheet!.Fonts!, _fonts, _fontIds);
    private uint FindOrCreateFill(Fill fill) => FindOrAppend(fill, _stylesheet!.Fills!, _fills, _fillIds);
    private uint FindOrCreateBorder(Border border) => FindOrAppend(border, _stylesheet!.Borders!, _borders, _borderIds);
    private uint FindOrCreateCellFormat(CellFormat format) => FindOrAppend(format, _stylesheet!.CellFormats!, _cellFormats, _cellFormatIds);

    private uint FindOrAppend<T>(T item, OpenXmlCompositeElement collection, List<T> items, Dictionary<string, uint> ids) where T : OpenXmlElement
    {
        var key = item.OuterXml;
        if (ids.TryGetValue(key, out var existing)) return existing;
        collection.Append(item);
        items.Add(item);
        collection.SetAttribute(new OpenXmlAttribute("count", string.Empty, items.Count.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        var id = checked((uint)items.Count - 1);
        ids[key] = id;
        _dirty = true;
        return id;
    }

    private uint DefaultFontId() => _cellFormats[0].FontId?.Value ?? 0;
    private uint DefaultFillId() => _cellFormats[0].FillId?.Value ?? 0;
    private uint DefaultBorderId() => _cellFormats[0].BorderId?.Value ?? 0;

    private void EnsureWritableStylesheet()
    {
        if (_stylesheet is not null)
        {
            if (_stylesheet.Fonts is null || _stylesheet.Fills is null || _stylesheet.Borders is null || _stylesheet.CellFormats is null ||
                _fonts.Count == 0 || _fills.Count < 2 || _borders.Count == 0 || _cellFormats.Count == 0)
                throw Invalid("Workbook stylesheet is missing required default style resources.");
            return;
        }
        _stylesPart = _workbookPart.AddNewPart<WorkbookStylesPart>();
        _stylesheet = CreateMinimalStylesheet();
        _stylesPart.Stylesheet = _stylesheet;
        IndexExistingStyles(_stylesheet);
        _dirty = true;
    }

    private void IndexExistingStyles(Stylesheet stylesheet)
    {
        _customFormatsById.Clear();
        _customFormatIds.Clear();
        foreach (var format in stylesheet.NumberingFormats?.Elements<NumberingFormat>() ?? [])
        {
            if (format.NumberFormatId?.HasValue != true || format.FormatCode?.HasValue != true) throw XlsxNumberFormatCodec.Invalid("Workbook contains an incomplete custom number-format definition.");
            var id = format.NumberFormatId.Value;
            var code = XlsxNumberFormatCodec.Canonicalize(format.FormatCode.Value, $"numFmt {id}");
            if (id < 164) throw XlsxNumberFormatCodec.Invalid($"Custom number format {id} uses the reserved built-in ID range.");
            if (!_customFormatsById.TryAdd(id, code)) throw XlsxNumberFormatCodec.Invalid($"Workbook defines custom number format {id} more than once.");
            _customFormatIds.TryAdd(code, id);
        }
        _fonts = stylesheet.Fonts?.Elements<Font>().ToList() ?? [];
        _fills = stylesheet.Fills?.Elements<Fill>().ToList() ?? [];
        _borders = stylesheet.Borders?.Elements<Border>().ToList() ?? [];
        _cellFormats = stylesheet.CellFormats?.Elements<CellFormat>().ToList() ?? [];
        _differentialFormats = stylesheet.DifferentialFormats?.Elements<DifferentialFormat>().ToList() ?? [];
        _fontIds = IndexByXml(_fonts);
        _fillIds = IndexByXml(_fills);
        _borderIds = IndexByXml(_borders);
        _cellFormatIds = IndexByXml(_cellFormats);
        _differentialFormatIds = IndexByXml(_differentialFormats);
        _originalNumberingFormats = stylesheet.NumberingFormats?.Elements<NumberingFormat>().Select(item => item.OuterXml).ToArray() ?? [];
        _originalFonts = _fonts.Select(item => item.OuterXml).ToArray();
        _originalFills = _fills.Select(item => item.OuterXml).ToArray();
        _originalBorders = _borders.Select(item => item.OuterXml).ToArray();
        _originalCellFormats = _cellFormats.Select(item => item.OuterXml).ToArray();
        _originalDifferentialFormats = _differentialFormats.Select(item => item.OuterXml).ToArray();
        _originalOtherChildren = stylesheet.ChildElements.Where(item => item is not NumberingFormats and not Fonts and not Fills and not Borders and not CellFormats and not DifferentialFormats).Select(item => item.OuterXml).ToArray();
    }

    private static Dictionary<string, uint> IndexByXml<T>(IReadOnlyList<T> items) where T : OpenXmlElement
    {
        var result = new Dictionary<string, uint>(StringComparer.Ordinal);
        for (var index = 0; index < items.Count; index++) result.TryAdd(items[index].OuterXml, checked((uint)index));
        return result;
    }

    private void AssertOriginalStylesPreserved()
    {
        var numberingFormats = _stylesheet!.NumberingFormats?.Elements<NumberingFormat>().Select(item => item.OuterXml).ToArray() ?? [];
        var fonts = _stylesheet.Fonts?.Elements<Font>().Select(item => item.OuterXml).ToArray() ?? [];
        var fills = _stylesheet.Fills?.Elements<Fill>().Select(item => item.OuterXml).ToArray() ?? [];
        var borders = _stylesheet.Borders?.Elements<Border>().Select(item => item.OuterXml).ToArray() ?? [];
        var cellFormats = _stylesheet.CellFormats?.Elements<CellFormat>().Select(item => item.OuterXml).ToArray() ?? [];
        var differentialFormats = _stylesheet.DifferentialFormats?.Elements<DifferentialFormat>().Select(item => item.OuterXml).ToArray() ?? [];
        var otherChildren = _stylesheet.ChildElements.Where(item => item is not NumberingFormats and not Fonts and not Fills and not Borders and not CellFormats and not DifferentialFormats).Select(item => item.OuterXml).ToArray();
        if (!numberingFormats.Take(_originalNumberingFormats.Length).SequenceEqual(_originalNumberingFormats, StringComparer.Ordinal) ||
            !fonts.Take(_originalFonts.Length).SequenceEqual(_originalFonts, StringComparer.Ordinal) ||
            !fills.Take(_originalFills.Length).SequenceEqual(_originalFills, StringComparer.Ordinal) ||
            !borders.Take(_originalBorders.Length).SequenceEqual(_originalBorders, StringComparer.Ordinal) ||
            !cellFormats.Take(_originalCellFormats.Length).SequenceEqual(_originalCellFormats, StringComparer.Ordinal) ||
            !differentialFormats.Take(_originalDifferentialFormats.Length).SequenceEqual(_originalDifferentialFormats, StringComparer.Ordinal) ||
            !otherChildren.SequenceEqual(_originalOtherChildren, StringComparer.Ordinal))
            throw new CodecException("style_preservation_failed", "Existing workbook style resources changed while applying a cell style.", "xl/styles.xml");
    }

    private static Stylesheet CreateMinimalStylesheet() => new(
        new Fonts(new Font(new FontSize { Val = 11D }, new FontName { Val = "Aptos" })) { Count = 1U },
        new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })) { Count = 2U },
        new Borders(new Border()) { Count = 1U },
        new CellStyleFormats(new CellFormat()) { Count = 1U },
        new CellFormats(new CellFormat()) { Count = 1U },
        new CellStyles(new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U }) { Count = 1U });

    private static void ValidateColor(SpreadsheetColor? color, string location, string component)
    {
        if (color is null) return;
        if (color.SourceCase == SpreadsheetColor.SourceOneofCase.None) throw Invalid($"Cell {location} {component} color has no source.");
        if (color.SourceCase == SpreadsheetColor.SourceOneofCase.Rgb && (color.Rgb.Length != 6 || !color.Rgb.All(Uri.IsHexDigit))) throw Invalid($"Cell {location} {component} color must be six-digit RGB.");
        if (color.SourceCase == SpreadsheetColor.SourceOneofCase.Theme && color.Theme > 11) throw Invalid($"Cell {location} {component} theme color index must be 0 through 11.");
        if (color.SourceCase == SpreadsheetColor.SourceOneofCase.Indexed && color.Indexed > 65) throw Invalid($"Cell {location} {component} indexed color must be 0 through 65.");
        if (color.SourceCase == SpreadsheetColor.SourceOneofCase.Automatic && !color.Automatic) throw Invalid($"Cell {location} {component} automatic color must be true.");
        if (color.HasTint && (!double.IsFinite(color.Tint) || color.Tint < -1 || color.Tint > 1)) throw Invalid($"Cell {location} {component} tint must be between -1 and 1.");
    }

    private static bool IsAutomaticPatternBackground(BackgroundColor color) =>
        color.Indexed?.Value == 64U && color.Rgb is null && color.Theme is null && color.Auto is null && color.Tint is null;

    private static IEnumerable<(string Name, SpreadsheetBorderEdgeStyle? Edge)> BorderEdges(SpreadsheetBorderStyle border)
    {
        yield return ("left", border.Left); yield return ("right", border.Right); yield return ("top", border.Top); yield return ("bottom", border.Bottom);
        yield return ("diagonal", border.Diagonal); yield return ("start", border.Start); yield return ("end", border.End);
        yield return ("horizontal", border.Horizontal); yield return ("vertical", border.Vertical);
    }

    private static bool ElementBoolean(BooleanPropertyType? element) => element is not null && (element.Val?.Value ?? true);
    private static bool HasStyle(CellStyleArtifact style) => style.Font is not null || style.Fill is not null || style.Border is not null || style.Alignment is not null || style.Protection is not null;
    private static bool HasBorder(SpreadsheetBorderStyle border) => BorderEdges(border).Any(item => item.Edge is not null) || border.HasDiagonalUp || border.HasDiagonalDown || border.HasOutline;
    private static bool HasAlignment(SpreadsheetAlignmentStyle alignment) => alignment.HasHorizontal || alignment.HasVertical || alignment.HasWrapText || alignment.HasTextRotation || alignment.HasIndent || alignment.HasShrinkToFit || alignment.HasReadingOrder;
    private static CodecException Invalid(string message) => new("invalid_cell_style", message, "xl/styles.xml");
    private static CodecException InvalidTableColor(string message) => new("invalid_worksheet_table", message, "xl/styles.xml");
}
