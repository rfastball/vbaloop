using System.Text;
using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class FormDefinitionTests
{
    private const string Spec = """
        {"name":"frmNew","caption":"입력","width":300,"height":200,"controls":[
          {"name":"fraGroup","type":"Frame","left":10,"top":10,"width":200,"height":100,"controls":[
            {"name":"txtName","type":"TextBox","text":"한글","left":5,"top":20,"width":100,"height":20}]}]}
        """;

    public FormDefinitionTests() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    [Fact]
    public void ParsesNestedDesignAndPlainCode()
    {
        var d = FormDefinition.Parse(Spec, "\uFEFFOption Explicit\n' 한글\n\n", 949);
        Assert.Equal("frmNew", d.Form.Name);
        Assert.Equal("한글", d.Form.Controls.Single().Controls.Single().Properties["Text"]);
        Assert.Equal(300f, d.Form.Properties["Width"]);
        Assert.Equal(new[] { "Option Explicit", "' 한글" }, d.Code);
    }

    [Theory]
    [InlineData("\"width\":300", "\"width\":0")]
    [InlineData("\"width\":300", "\"width\":\"300\"")]
    [InlineData("\"width\":300", "\"width\":1e999")]
    [InlineData("\"width\":300", "\"width\":300,\"width\":301")]
    [InlineData("\"caption\":", "\"captoin\":")]
    [InlineData("\"TextBox\"", "\"Unknown.ActiveX\"")]
    [InlineData("\"txtName\"", "\"FRAGROUP\"")]
    [InlineData("\"frmNew\"", "\"../escape\"")]
    [InlineData("\"frmNew\"", "\"CON\"")]
    [InlineData("\"한글\"", "\"😀\"")]
    public void RejectsInvalidDesign(string from, string to)
    {
        Assert.Throws<CliError>(() => FormDefinition.Parse(Spec.Replace(from, to), "", 949));
    }

    [Theory]
    [InlineData("Attribute VB_Name = \"Other\"")]
    [InlineData("VERSION 5.00")]
    [InlineData("Option Explicit\n' 😀")]
    [InlineData("Option Explicit\0")]
    public void RejectsExportHeadersAndLossyCode(string code) =>
        Assert.Throws<CliError>(() => FormDefinition.Parse(Spec, code, 949));
}
