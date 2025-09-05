//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelFont : IExcelFont
{
    private MsExcel.Font _font;
    private bool _disposedValue;
    public string Name
    {
        get => _font.Name?.ToString();
        set => _font.Name = value;
    }

    public double Size
    {
        get => Convert.ToDouble(_font.Size);
        set => _font.Size = value;
    }

    public bool Bold
    {
        get => Convert.ToBoolean(_font.Bold);
        set => _font.Bold = value;
    }

    public bool Strikethrough
    {
        get => Convert.ToBoolean(_font.Strikethrough);
        set => _font.Bold = value;
    }

    public bool Italic
    {
        get => Convert.ToBoolean(_font.Italic);
        set => _font.Italic = value;
    }

    public int ColorIndex
    {
        get => (int)_font.ColorIndex;
        set => _font.ColorIndex = value;
    }

    public object FontStyle
    {
        get => _font.FontStyle;
        set => _font.FontStyle = value;
    }

    public Color Color
    {
        get => ColorTranslator.FromOle(Convert.ToInt32(_font.Color));
        set => _font.Color = ColorTranslator.ToOle(value);
    }

    public XlUnderlineStyle Underline
    {
        get => (XlUnderlineStyle)_font.Underline;
        set => _font.Underline = (MsExcel.XlUnderlineStyle)value;
    }

    public bool Superscript
    {
        get => Convert.ToBoolean(_font.Superscript);
        set => _font.Superscript = value;
    }

    public bool Subscript
    {
        get => Convert.ToBoolean(_font.Subscript);
        set => _font.Subscript = value;
    }

    internal ExcelFont(MsExcel.Font font)
    {
        _font = font;
        _disposedValue = false;
    }


    private void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _font != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_font) > 0) { }
            }
            catch { }
            _font = null;
        }

        _disposedValue = true;
    }

    public void Dispose() => Dispose(true);
}
