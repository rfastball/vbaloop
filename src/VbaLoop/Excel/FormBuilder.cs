using System.Reflection;
using VbaLoop.Core;

namespace VbaLoop.Excel;

public static class FormBuilder
{
    public static void Export(dynamic project, FormDefinition definition, string directory, int codePage)
    {
        var node = definition.Form;
        dynamic form = project.VBComponents.Add(VbaProjectOps.TypeMsForm);
        form.Name = node.Name;
        foreach (var (key, value) in node.Properties) form.Properties.Item(key).Value = value;
        AddControls(form.Designer, node.Controls);
        VbaProjectOps.ReplaceCode(form.CodeModule, definition.Code);
        string path = Path.Combine(directory, node.Name + ".frm");
        form.Export(path);
        VbaProjectOps.VerifiedRemove(project, form, node.Name);
        dynamic imported = project.VBComponents.Import(path);
        if ((string)imported.Name != node.Name || (int)imported.Type != VbaProjectOps.TypeMsForm)
            throw new CliError("generated form changed identity on Import");
        dynamic designer = imported.Designer; // Materialize the imported designer before reading Properties.
        foreach (var (key, value) in node.Properties) Verify(imported.Properties.Item(key).Value, value, node.Name + "." + key, 1);
        VerifyControls(designer, node.Controls);
        string actual = VbaProjectOps.ReadModuleBody(imported.CodeModule);
        var actualLines = CanonicalText.TrimTrailingEmpty(CanonicalText.SplitLines(actual.TrimStart('\r', '\n')));
        var expected = definition.Code.SkipWhile(string.IsNullOrWhiteSpace).ToArray();
        if (!VbaCasing.ReconcileBlock(actualLines, expected).SequenceEqual(expected))
            throw new CliError("generated form code did not survive Export/Import");
        string exported = Cp949Guard.GetStrict(codePage).GetString(File.ReadAllBytes(path));
        CanonicalText.WriteCanonical(path, CanonicalText.StripMemberAttributes(
            CanonicalText.TrimTrailingEmpty(CanonicalText.SplitLines(exported))));
    }

    private static void AddControls(dynamic container, IReadOnlyList<FormDefinition.Node> nodes)
    {
        foreach (var node in nodes)
        {
            dynamic control = container.Controls.Add($"Forms.{node.Type}.1", node.Name, true);
            foreach (var (key, value) in node.Properties)
                ((object)control).GetType().InvokeMember(key, BindingFlags.SetProperty, null, control, new[] { value });
            AddControls(control, node.Controls);
        }
    }

    private static void VerifyControls(dynamic container, IReadOnlyList<FormDefinition.Node> nodes)
    {
        foreach (var node in nodes)
        {
            dynamic control = container.Controls.Item(node.Name);
            foreach (var (key, value) in node.Properties)
            {
                object actual = ((object)control).GetType().InvokeMember(key, BindingFlags.GetProperty, null, control, null)!;
                Verify(actual, value, node.Name + "." + key, 0.1);
            }
            VerifyControls(control, node.Controls);
        }
    }

    private static void Verify(object actual, object expected, string property, double tolerance)
    {
        bool equal = expected is string s ? Convert.ToString(actual) == s : Math.Abs(Convert.ToDouble(actual) - Convert.ToDouble(expected)) <= tolerance;
        if (!equal) throw new CliError($"generated form {property}: expected '{expected}', read back '{actual}'");
    }
}
