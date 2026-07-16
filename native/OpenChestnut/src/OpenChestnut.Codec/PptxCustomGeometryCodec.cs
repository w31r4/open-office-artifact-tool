using System.Globalization;
using DocumentFormat.OpenXml;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

// Bounded literal DrawingML custom paths used by source-built presentation
// templates. Guides, handles, connection sites, text rectangles, arcs,
// quadratic curves, and per-path paint overrides stay outside this slice.
internal static class PptxCustomGeometryCodec
{
    private const int MaxPaths = 64;
    private const int MaxCommands = 16_384;
    private const long MaxCoordinate = int.MaxValue;

    internal static bool Supports(A.CustomGeometry? geometry)
    {
        if (geometry is null || geometry.ChildElements.Count != 2 ||
            geometry.ChildElements[0] is not A.AdjustValueList adjustValues || adjustValues.HasChildren || adjustValues.HasAttributes ||
            geometry.ChildElements[1] is not A.PathList pathList || pathList.HasAttributes)
            return false;
        var paths = pathList.Elements<A.Path>().ToArray();
        if (paths.Length is < 1 or > MaxPaths || pathList.ChildElements.Count != paths.Length) return false;
        var commandCount = 0;
        return paths.All(path => Supports(path, ref commandCount));
    }

    internal static IEnumerable<PresentationCustomGeometryPath> Read(A.CustomGeometry? geometry)
    {
        if (!Supports(geometry)) yield break;
        foreach (var nativePath in geometry!.GetFirstChild<A.PathList>()!.Elements<A.Path>())
        {
            var path = new PresentationCustomGeometryPath
            {
                Width = checked((long)nativePath.Width!.Value),
                Height = checked((long)nativePath.Height!.Value),
            };
            foreach (var nativeCommand in nativePath.ChildElements)
            {
                var command = nativeCommand switch
                {
                    A.MoveTo move => new PresentationCustomGeometryCommand { MoveTo = ReadPoint(move.Point!) },
                    A.LineTo line => new PresentationCustomGeometryCommand { LineTo = ReadPoint(line.Point!) },
                    A.CubicBezierCurveTo cubic => ReadCubic(cubic),
                    A.CloseShapePath => new PresentationCustomGeometryCommand { Close = true },
                    _ => throw new InvalidOperationException("Unsupported custom geometry command passed the recognition gate."),
                };
                path.Commands.Add(command);
            }
            yield return path;
        }
    }

    internal static void Validate(PresentationShape shape, string shapeId)
    {
        if (shape.Geometry != "custom")
        {
            if (shape.CustomPaths.Count > 0)
                throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has custom paths without custom geometry.");
            return;
        }
        if (shape.CustomPaths.Count is < 1 or > MaxPaths)
            throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} custom geometry must contain 1 through {MaxPaths} paths.");
        var commandCount = 0;
        foreach (var path in shape.CustomPaths)
        {
            if (path.Width is <= 0 or > MaxCoordinate || path.Height is <= 0 or > MaxCoordinate || path.Commands.Count == 0)
                throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an invalid custom path extent or empty command list.");
            commandCount += path.Commands.Count;
            if (commandCount > MaxCommands)
                throw new CodecException("presentation_item_budget_exceeded", $"Presentation shape {shapeId} custom geometry exceeds the {MaxCommands}-command budget.");
            foreach (var command in path.Commands) Validate(command, shapeId);
        }
    }

    internal static void Apply(P.ShapeProperties properties, PresentationShape shape)
    {
        if (shape.Geometry != "custom")
        {
            properties.GetFirstChild<A.CustomGeometry>()?.Remove();
            var preset = properties.GetFirstChild<A.PresetGeometry>();
            if (preset is null)
            {
                preset = new A.PresetGeometry(new A.AdjustValueList());
                var presetTransform = properties.GetFirstChild<A.Transform2D>();
                if (presetTransform is null) properties.PrependChild(preset);
                else properties.InsertAfter(preset, presetTransform);
            }
            preset.Preset = shape.Geometry switch
            {
                "ellipse" => A.ShapeTypeValues.Ellipse,
                "roundRect" => A.ShapeTypeValues.RoundRectangle,
                _ => A.ShapeTypeValues.Rectangle,
            };
            return;
        }
        properties.GetFirstChild<A.PresetGeometry>()?.Remove();
        properties.GetFirstChild<A.CustomGeometry>()?.Remove();
        OpenXmlElement geometry = Build(shape);
        var transform = properties.GetFirstChild<A.Transform2D>();
        if (transform is null) properties.PrependChild(geometry);
        else properties.InsertAfter(geometry, transform);
    }

    private static A.CustomGeometry Build(PresentationShape shape)
    {
        var paths = new A.PathList();
        foreach (var source in shape.CustomPaths)
        {
            var path = new A.Path { Width = source.Width, Height = source.Height };
            foreach (var command in source.Commands)
            {
                path.Append(command.CommandCase switch
                {
                    PresentationCustomGeometryCommand.CommandOneofCase.MoveTo => new A.MoveTo(Point(command.MoveTo)),
                    PresentationCustomGeometryCommand.CommandOneofCase.LineTo => new A.LineTo(Point(command.LineTo)),
                    PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo => new A.CubicBezierCurveTo(
                        Point(command.CubicBezierTo.Control1),
                        Point(command.CubicBezierTo.Control2),
                        Point(command.CubicBezierTo.End)),
                    PresentationCustomGeometryCommand.CommandOneofCase.Close => new A.CloseShapePath(),
                    _ => throw new CodecException("invalid_presentation_geometry", "Presentation custom geometry contains an empty command."),
                });
            }
            paths.Append(path);
        }
        return new A.CustomGeometry(new A.AdjustValueList(), paths);
    }

    private static bool Supports(A.Path path, ref int commandCount)
    {
        if (path.Width?.Value is not { } width || width is 0 or > MaxCoordinate ||
            path.Height?.Value is not { } height || height is 0 or > MaxCoordinate ||
            !HasOnlyAttributes(path, "w", "h") || path.ChildElements.Count == 0)
            return false;
        commandCount += path.ChildElements.Count;
        if (commandCount > MaxCommands) return false;
        return path.ChildElements.All(command => command switch
        {
            A.MoveTo move => SupportsPointContainer(move, move.Point, 1),
            A.LineTo line => SupportsPointContainer(line, line.Point, 1),
            A.CubicBezierCurveTo cubic => !cubic.HasAttributes && cubic.ChildElements.Count == 3 && cubic.Elements<A.Point>().All(SupportsPoint),
            A.CloseShapePath close => !close.HasAttributes && !close.HasChildren,
            _ => false,
        });
    }

    private static bool SupportsPointContainer(OpenXmlCompositeElement container, A.Point? point, int childCount) =>
        !container.HasAttributes && container.ChildElements.Count == childCount && point is not null && SupportsPoint(point);

    private static bool SupportsPoint(A.Point point) =>
        !point.HasChildren && HasOnlyAttributes(point, "x", "y") &&
        TryCoordinate(point.X?.Value, out _) && TryCoordinate(point.Y?.Value, out _);

    private static PresentationCustomGeometryPoint ReadPoint(A.Point point) => new()
    {
        X = ParseCoordinate(point.X!.Value!),
        Y = ParseCoordinate(point.Y!.Value!),
    };

    private static PresentationCustomGeometryCommand ReadCubic(A.CubicBezierCurveTo source)
    {
        var points = source.Elements<A.Point>().ToArray();
        return new PresentationCustomGeometryCommand
        {
            CubicBezierTo = new PresentationCustomGeometryCubicBezier
            {
                Control1 = ReadPoint(points[0]),
                Control2 = ReadPoint(points[1]),
                End = ReadPoint(points[2]),
            },
        };
    }

    private static A.Point Point(PresentationCustomGeometryPoint source) => new()
    {
        X = source.X.ToString(CultureInfo.InvariantCulture),
        Y = source.Y.ToString(CultureInfo.InvariantCulture),
    };

    private static void Validate(PresentationCustomGeometryCommand command, string shapeId)
    {
        switch (command.CommandCase)
        {
            case PresentationCustomGeometryCommand.CommandOneofCase.MoveTo:
                Validate(command.MoveTo, shapeId);
                break;
            case PresentationCustomGeometryCommand.CommandOneofCase.LineTo:
                Validate(command.LineTo, shapeId);
                break;
            case PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo:
                if (command.CubicBezierTo.Control1 is null || command.CubicBezierTo.Control2 is null || command.CubicBezierTo.End is null)
                    throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an incomplete cubic Bézier command.");
                Validate(command.CubicBezierTo.Control1, shapeId);
                Validate(command.CubicBezierTo.Control2, shapeId);
                Validate(command.CubicBezierTo.End, shapeId);
                break;
            case PresentationCustomGeometryCommand.CommandOneofCase.Close:
                if (!command.Close) throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has an invalid close command.");
                break;
            default:
                throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} contains an empty custom geometry command.");
        }
    }

    private static void Validate(PresentationCustomGeometryPoint? point, string shapeId)
    {
        if (point is null || point.X < -MaxCoordinate || point.X > MaxCoordinate || point.Y < -MaxCoordinate || point.Y > MaxCoordinate)
            throw new CodecException("invalid_presentation_geometry", $"Presentation shape {shapeId} has a custom path point outside the signed 32-bit coordinate range.");
    }

    private static bool TryCoordinate(string? value, out long coordinate) =>
        long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out coordinate) &&
        coordinate >= -MaxCoordinate && coordinate <= MaxCoordinate;

    private static long ParseCoordinate(string value) => long.Parse(value, NumberStyles.Integer, CultureInfo.InvariantCulture);

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute => allowed.Contains(attribute.LocalName));
    }
}
