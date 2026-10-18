using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace VbaLoop.Core;

/// <summary>Agent-authored JSON design and plain VBA body. No arbitrary ActiveX or property dispatch.</summary>
public sealed record FormDefinition(FormDefinition.Node Form, IReadOnlyList<string> Code)
{
    public sealed record Node(string Name, string Type, IReadOnlyDictionary<string, object> Properties, IReadOnlyList<Node> Controls);

    public static FormDefinition Parse(string json, string code, int codePage)
    {
        try
        {
            using var doc = JsonDocument.Parse(json.TrimStart('\uFEFF'));
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var encoding = Cp949Guard.GetStrict(codePage);
            Node Read(JsonElement element, bool root)
            {
                if (element.ValueKind != JsonValueKind.Object) throw new CliError("form/control must be a JSON object");
                string Text(string key)
                {
                    if (!element.TryGetProperty(key, out var value) || value.ValueKind != JsonValueKind.String)
                        throw new CliError($"form/control requires string '{key}'");
                    string text = value.GetString()!;
                    encoding.GetByteCount(text);
                    if (text.Contains('\0')) throw new CliError($"'{key}' contains a NUL character");
                    return text;
                }
                string name = Text("name"), type = root ? "UserForm" : Text("type");
                if (!Regex.IsMatch(name, @"\A[A-Za-z][A-Za-z0-9_]{0,39}\z") ||
                    Regex.IsMatch(name, @"\A(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])\z", RegexOptions.IgnoreCase) || !names.Add(name))
                    throw new CliError($"invalid or duplicate form/control name: {name}");
                if (!root && type is not ("Label" or "TextBox" or "CommandButton" or "Frame"))
                    throw new CliError($"{name}: unsupported control type '{type}' (Label, TextBox, CommandButton, Frame only)");
                var allowed = new HashSet<string>(StringComparer.Ordinal) { "name", "width", "height" };
                if (!root) allowed.UnionWith(new[] { "type", "left", "top" });
                allowed.Add(type == "TextBox" ? "text" : "caption");
                if (root || type == "Frame") allowed.Add("controls");
                var seen = new HashSet<string>(StringComparer.Ordinal);
                foreach (var property in element.EnumerateObject())
                    if (!allowed.Contains(property.Name) || !seen.Add(property.Name))
                        throw new CliError($"{name}: unknown or duplicate property '{property.Name}'");
                var props = new Dictionary<string, object>();
                foreach (string key in root ? new[] { "width", "height" } : new[] { "left", "top", "width", "height" })
                {
                    if (!element.TryGetProperty(key, out var value) || value.ValueKind != JsonValueKind.Number ||
                        !value.TryGetDouble(out double number) || !double.IsFinite(number) || number < 0 || number > 10000 ||
                        (key is "width" or "height" && number == 0))
                        throw new CliError($"{name}.{key}: expected {(key is "width" or "height" ? "positive" : "non-negative")} points, at most 10000");
                    props[char.ToUpperInvariant(key[0]) + key[1..]] = (float)number;
                }
                string textKey = type == "TextBox" ? "text" : "caption";
                if (element.TryGetProperty(textKey, out _)) props[textKey == "text" ? "Text" : "Caption"] = Text(textKey);
                var children = new List<Node>();
                if (element.TryGetProperty("controls", out var controls))
                {
                    if (controls.ValueKind != JsonValueKind.Array) throw new CliError($"{name}.controls: expected an array");
                    foreach (var child in controls.EnumerateArray()) children.Add(Read(child, false));
                }
                return new Node(name, type, props, children);
            }
            var form = Read(doc.RootElement, true);
            var lines = CanonicalText.TrimTrailingEmpty(CanonicalText.SplitLines(code.TrimStart('\uFEFF')));
            for (int i = 0; i < lines.Length; i++)
                if (Regex.IsMatch(lines[i], @"^\s*(Attribute\s|VERSION\s|Begin\s+\{)", RegexOptions.IgnoreCase) || lines[i].Contains('\0'))
                    throw new CliError($"code:{i + 1}: use a plain VBA body without export headers, Attributes or NUL characters");
            var violations = Cp949Guard.Check(lines, codePage);
            if (violations.Count > 0)
                throw new CliError($"code:{violations[0].Line}:{violations[0].Column}: {violations[0].CodePoint} cannot be encoded in CP{codePage}");
            return new FormDefinition(form, lines);
        }
        catch (JsonException ex) { throw new CliError($"invalid form JSON: {ex.Message}"); }
        catch (EncoderFallbackException) { throw new CliError($"form text cannot be encoded in CP{codePage}"); }
    }
}
