//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel TextFrame 对象的二次封装实现类
/// </summary>
internal class ExcelTextFrame : IExcelTextFrame
{
    private MsExcel.TextFrame _textFrame;
    private bool _disposedValue;

    internal ExcelTextFrame(MsExcel.TextFrame textFrame)
    {
        _textFrame = textFrame;
        _disposedValue = false;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                if (_textFrame != null)
                    Marshal.ReleaseComObject(_textFrame);
            }
            catch { }
            _textFrame = null;
        }

        _disposedValue = true;
    }

    public void Dispose() => Dispose(true);

    public MsoTextOrientation Orientation
    {
        get => _textFrame != null ? (MsoTextOrientation)_textFrame.Orientation : 0;
        set
        {
            if (_textFrame != null)
                _textFrame.Orientation = (MsCore.MsoTextOrientation)value;
        }
    }

    public bool AutoSize
    {
        get => _textFrame != null && Convert.ToBoolean(_textFrame.AutoSize);
        set { if (_textFrame != null) _textFrame.AutoSize = value; }
    }

    public XlHAlign HorizontalAlignment
    {
        get => _textFrame != null ? (XlHAlign)_textFrame.HorizontalAlignment : XlHAlign.xlHAlignLeft;
        set { if (_textFrame != null) _textFrame.HorizontalAlignment = (MsExcel.XlHAlign)value; }
    }

    public XlVAlign VerticalAlignment
    {
        get => _textFrame != null ? (XlVAlign)_textFrame.VerticalAlignment : XlVAlign.xlVAlignJustify;
        set
        {
            if (_textFrame != null)
                _textFrame.VerticalAlignment = (MsExcel.XlVAlign)value;
        }
    }

    public float MarginLeft
    {
        get => _textFrame?.MarginLeft ?? 0;
        set { if (_textFrame != null) _textFrame.MarginLeft = value; }
    }

    public float MarginRight
    {
        get => _textFrame?.MarginRight ?? 0;
        set { if (_textFrame != null) _textFrame.MarginRight = value; }
    }

    public float MarginTop
    {
        get => _textFrame?.MarginTop ?? 0;
        set { if (_textFrame != null) _textFrame.MarginTop = value; }
    }

    public float MarginBottom
    {
        get => _textFrame?.MarginBottom ?? 0;
        set { if (_textFrame != null) _textFrame.MarginBottom = value; }
    }

    public IExcelCharacters? Characters(int? start = null, int? length = null)
    {
        var charactersObj = _textFrame?.Characters(start.ComArgsVal(), length.ComArgsVal());
        MsExcel.Range? range = null;
        if (_textFrame?.Parent is MsExcel.Range r1)
            range = r1;
        if (_textFrame?.Parent is MsExcel.Shape shape)
            if (shape.Parent is MsExcel.Range r2)
                range = r2;
        if (_textFrame?.Parent is MsExcel.ShapeRange shapeRange)
            if (shapeRange.Parent is MsExcel.Range r3)
                range = r3;
        return charactersObj != null ? new ExcelCharacters(charactersObj, range) : null;
    }
}