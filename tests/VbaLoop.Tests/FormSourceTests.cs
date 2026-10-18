using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class FormSourceTests
{
    [Fact]
    public void PublishRequiresForceAndRestoresFirstFileWhenSecondMoveFails()
    {
        string root = Path.Combine(Path.GetTempPath(), "vbaloop-test", Guid.NewGuid().ToString("N"));
        string src = Path.Combine(root, "src"), stage = Path.Combine(root, "stage");
        Directory.CreateDirectory(src);
        Directory.CreateDirectory(stage);
        try
        {
            void Stage()
            {
                File.WriteAllText(Path.Combine(stage, "Form.frm"), "new frm");
                File.WriteAllText(Path.Combine(stage, "Form.frx"), "new frx");
            }
            Stage();
            FormSource.Publish(stage, src, "Form", false);
            Assert.Equal("new frx", File.ReadAllText(Path.Combine(src, "Form.frx")));
            File.WriteAllText(Path.Combine(src, "Form.frm"), "old frm");
            Stage();
            Assert.Throws<CliError>(() => FormSource.Publish(stage, src, "Form", false));
            Assert.Equal("old frm", File.ReadAllText(Path.Combine(src, "Form.frm")));
            File.Delete(Path.Combine(src, "Form.frx"));
            Directory.CreateDirectory(Path.Combine(src, "Form.frx"));
            Assert.Throws<CliError>(() => FormSource.Publish(stage, src, "Form", true));
            Assert.Equal("old frm", File.ReadAllText(Path.Combine(src, "Form.frm")));
            Assert.True(Directory.Exists(Path.Combine(src, "Form.frx")));
        }
        finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void MissingPairAndOtherComponentTypeNeverReplaceSource()
    {
        string root = Path.Combine(Path.GetTempPath(), "vbaloop-test", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            File.WriteAllText(Path.Combine(root, "Form.bas"), "keep");
            Assert.Throws<CliError>(() => FormSource.CheckDestination(root, "Form", true));
            Assert.Throws<CliError>(() => FormSource.Publish(root, Path.Combine(root, "src"), "Missing", false));
            Assert.False(Directory.Exists(Path.Combine(root, "src")));
            Assert.Equal("keep", File.ReadAllText(Path.Combine(root, "Form.bas")));
        }
        finally { Directory.Delete(root, true); }
    }
}
