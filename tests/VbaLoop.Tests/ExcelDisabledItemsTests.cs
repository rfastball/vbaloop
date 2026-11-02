using System.Text;
using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class ExcelDisabledItemsTests
{
    private static byte[] Entry(int type, string path, int? lenOverride = null)
    {
        byte[] p = Encoding.Unicode.GetBytes(path + "\0");
        var ms = new MemoryStream();
        var w = new BinaryWriter(ms);
        w.Write(type);
        w.Write(lenOverride ?? p.Length);
        w.Write(4);
        w.Write(p);
        w.Write(0);
        return ms.ToArray();
    }

    [Fact]
    public void DecodesDocumentEntry()
    {
        // 실측 값(144바이트)과 같은 배치
        const string path = @"c:\users\rfast\desktop\vba_projects\지출 리팩토링\지출원인행위정보 송신 목록.xlsm";
        byte[] v = Entry(2, path);
        Assert.Equal(144, v.Length);
        Assert.Equal(path, ExcelDisabledItems.DecodeDocumentPath(v));
    }

    [Fact]
    public void RejectsNonDocumentOrMalformed()
    {
        Assert.Null(ExcelDisabledItems.DecodeDocumentPath(Entry(1, @"c:\addin.xlam")));
        Assert.Null(ExcelDisabledItems.DecodeDocumentPath(Entry(2, @"c:\a.xlsm", lenOverride: 9999)));
        Assert.Null(ExcelDisabledItems.DecodeDocumentPath(Entry(2, @"c:\a.xlsm", lenOverride: 3)));
        Assert.Null(ExcelDisabledItems.DecodeDocumentPath(new byte[8]));
    }
}
