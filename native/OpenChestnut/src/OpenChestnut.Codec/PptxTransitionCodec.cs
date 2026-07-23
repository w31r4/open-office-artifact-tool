using System.Security.Cryptography;
using System.Text;
using System.Globalization;
using DocumentFormat.OpenXml;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenChestnut.Codec;

// Owns one direct p:transition leaf on one SlidePart. This deliberately does
// not model p:timing, sound actions, p14 duration, or any of the large native
// effect vocabulary. The bounded profile keeps slideshow advancement usable
// without pretending that it is a general animation engine.
internal static class PptxTransitionCodec
{
    internal const uint MaxAdvanceAfterMilliseconds = 86_400_000;

    internal static PresentationTransition? Read(P.Slide source)
    {
        var transitions = source.Elements<P.Transition>().ToArray();
        return transitions.Length == 1 && TryRead(transitions[0], out var semantic) ? semantic : null;
    }

    internal static bool HasTransition(P.Slide source) => source.Elements<P.Transition>().Any();

    internal static bool Supports(P.Slide source)
    {
        var transitions = source.Elements<P.Transition>().ToArray();
        return transitions.Length == 1 && TryRead(transitions[0], out _);
    }

    internal static void Validate(PresentationTransition? source)
    {
        if (source is null) return;
        if (!source.HasAdvanceOnClick)
            throw Invalid("Presentation transition requires an explicit advance_on_click value.");
        if (source.HasAdvanceAfterMs && source.AdvanceAfterMs > MaxAdvanceAfterMilliseconds)
            throw Invalid($"Presentation transition advance_after_ms must not exceed {MaxAdvanceAfterMilliseconds}.");
        if (!IsSpeed(source.Speed))
            throw Invalid("Presentation transition speed must be slow, medium, or fast.");
        switch (source.Effect)
        {
            case "fade" when string.IsNullOrEmpty(source.Direction):
                return;
            case "push" when IsDirection(source.Direction):
                return;
            case "fade":
                throw Invalid("Presentation fade transition must not carry direction.");
            case "push":
                throw Invalid("Presentation push transition direction must be left, up, right, or down.");
            default:
                throw Invalid("Presentation transition effect must be fade or push.");
        }
    }

    internal static void Build(P.Slide target, PresentationTransition? source)
    {
        if (source is null) return;
        target.AddChild(BuildElement(source), true);
    }

    internal static void Apply(P.Slide target, PresentationTransition? source)
    {
        Validate(source);
        var transitions = target.Elements<P.Transition>().ToArray();
        if (transitions.Length > 1)
            throw new CodecException("presentation_transition_topology_changed", "Slide contains multiple transition elements.");
        var current = transitions.SingleOrDefault();
        if (source is null)
        {
            current?.Remove();
            return;
        }
        var replacement = BuildElement(source);
        if (current is null)
        {
            target.AddChild(replacement, true);
            return;
        }
        current.InsertAfterSelf(replacement);
        current.Remove();
    }

    internal static string SemanticHash(PresentationTransition? source) =>
        Hash((source?.Clone() ?? new PresentationTransition()).ToByteArray());

    // Used only for a no-op proof. Semantic hashing intentionally maps opaque
    // transitions to absence; this raw hash makes sure that an unsupported
    // native transition was not silently dropped or rewritten.
    internal static string ElementHash(P.Slide source)
    {
        var xml = string.Concat(source.Elements<P.Transition>().Select(transition => transition.OuterXml));
        return Hash(Encoding.UTF8.GetBytes(xml));
    }

    private static bool TryRead(P.Transition source, out PresentationTransition semantic)
    {
        semantic = new PresentationTransition();
        if (source.ExtendedAttributes.Any() || !HasTransitionAttributes(source) ||
            source.Speed?.Value is not { } speed || source.AdvanceOnClick?.Value is not { } advanceOnClick ||
            !TrySpeed(speed, out var speedName) || source.ChildElements.Count != 1)
            return false;
        semantic.Speed = speedName;
        semantic.AdvanceOnClick = advanceOnClick;
        if (source.AdvanceAfterTime?.Value is { } advanceAfterText)
        {
            if (!uint.TryParse(advanceAfterText, NumberStyles.None, CultureInfo.InvariantCulture, out var advanceAfter)) return false;
            if (advanceAfter > MaxAdvanceAfterMilliseconds) return false;
            semantic.AdvanceAfterMs = advanceAfter;
        }
        switch (source.FirstChild)
        {
            case P.FadeTransition fade when IsEmpty(fade):
                semantic.Effect = "fade";
                return true;
            case P.PushTransition push when IsPush(push, out var direction):
                semantic.Effect = "push";
                semantic.Direction = direction;
                return true;
            default:
                return false;
        }
    }

    private static P.Transition BuildElement(PresentationTransition source)
    {
        Validate(source);
        var transition = new P.Transition
        {
            Speed = Speed(source.Speed),
            AdvanceOnClick = source.AdvanceOnClick,
        };
        if (source.HasAdvanceAfterMs) transition.AdvanceAfterTime = source.AdvanceAfterMs.ToString(CultureInfo.InvariantCulture);
        transition.Append(source.Effect switch
        {
            "fade" => new P.FadeTransition(),
            "push" => new P.PushTransition { Direction = Direction(source.Direction) },
            _ => throw Invalid("Presentation transition effect is invalid."),
        });
        return transition;
    }

    private static bool IsEmpty(OpenXmlElement source) =>
        !source.ExtendedAttributes.Any() && source.GetAttributes().Count == 0 && source.ChildElements.Count == 0;

    private static bool IsPush(P.PushTransition source, out string direction)
    {
        direction = string.Empty;
        if (source.ExtendedAttributes.Any() || !HasOnlyAttributes(source, "dir") || source.ChildElements.Count != 0 || source.Direction?.Value is not { } value)
            return false;
        return TryDirection(value, out direction);
    }

    private static bool HasOnlyAttributes(OpenXmlElement source, params string[] names)
    {
        var attributes = source.GetAttributes();
        if (attributes.Count != names.Length) return false;
        var expected = names.ToHashSet(StringComparer.Ordinal);
        return attributes.All(attribute => attribute.NamespaceUri.Length == 0 && expected.Remove(attribute.LocalName)) && expected.Count == 0;
    }

    private static bool HasTransitionAttributes(P.Transition source)
    {
        var attributes = source.GetAttributes();
        if (attributes.Count is < 2 or > 3) return false;
        var allowed = new HashSet<string>(["spd", "advClick", "advTm"], StringComparer.Ordinal);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var attribute in attributes)
        {
            if (attribute.NamespaceUri.Length != 0 || !allowed.Contains(attribute.LocalName) || !seen.Add(attribute.LocalName)) return false;
        }
        return seen.Contains("spd") && seen.Contains("advClick");
    }

    private static bool TrySpeed(P.TransitionSpeedValues value, out string name)
    {
        name = value.Equals(P.TransitionSpeedValues.Slow)
            ? "slow"
            : value.Equals(P.TransitionSpeedValues.Medium)
                ? "medium"
                : value.Equals(P.TransitionSpeedValues.Fast)
                    ? "fast"
                    : string.Empty;
        return name.Length > 0;
    }

    private static P.TransitionSpeedValues Speed(string value) => value switch
    {
        "slow" => P.TransitionSpeedValues.Slow,
        "medium" => P.TransitionSpeedValues.Medium,
        "fast" => P.TransitionSpeedValues.Fast,
        _ => throw Invalid("Presentation transition speed is invalid."),
    };

    private static bool TryDirection(P.TransitionSlideDirectionValues value, out string name)
    {
        name = value.Equals(P.TransitionSlideDirectionValues.Left)
            ? "left"
            : value.Equals(P.TransitionSlideDirectionValues.Up)
                ? "up"
                : value.Equals(P.TransitionSlideDirectionValues.Right)
                    ? "right"
                    : value.Equals(P.TransitionSlideDirectionValues.Down)
                        ? "down"
                        : string.Empty;
        return name.Length > 0;
    }

    private static P.TransitionSlideDirectionValues Direction(string value) => value switch
    {
        "left" => P.TransitionSlideDirectionValues.Left,
        "up" => P.TransitionSlideDirectionValues.Up,
        "right" => P.TransitionSlideDirectionValues.Right,
        "down" => P.TransitionSlideDirectionValues.Down,
        _ => throw Invalid("Presentation push transition direction is invalid."),
    };

    private static bool IsSpeed(string value) => value is "slow" or "medium" or "fast";
    private static bool IsDirection(string value) => value is "left" or "up" or "right" or "down";
    private static CodecException Invalid(string message) => new("invalid_presentation_transition", message);
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
