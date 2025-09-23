//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelDrawing : IExcelDrawing
{
    private MsExcel.Drawing? _drawing;
    private bool _disposedValue;

    /// <summary>
    /// 获取图表所在的 Excel Application 对象
    /// </summary>
    public IExcelApplication Application => new ExcelApplication(_drawing.Application);

    /// <summary>
    /// 获取图表对象所在的父对象
    /// </summary>
    public object? Parent
    {
        get
        {
            if (_drawing?.Parent == null)
            {
                return null;
            }
            if (_drawing.Parent is MsExcel.DrawingObjects chartObjs)
            {
                return new ExcelDrawingObjects(chartObjs);
            }
            if (_drawing.Parent is MsExcel.Workbook workbook)
            {
                return new ExcelWorkbook(workbook);
            }
            if (_drawing.Parent is MsExcel.Worksheet worksheet)
            {
                return new ExcelWorksheet(worksheet);
            }
            return null;
        }
    }

    /// <summary>
    /// 获取绘图对象的索引
    /// </summary>
    public int Index => _drawing != null ? _drawing.Index : -1;

    public string Name
    {
        get => _drawing != null ? _drawing.Name : string.Empty;
        set
        {
            if (_drawing != null)
                _drawing.Name = value;
        }
    }

    public double Left
    {
        get => _drawing != null ? _drawing.Left : 0;
        set
        {
            if (_drawing != null)
                _drawing.Left = value;
        }
    }

    public double Top
    {
        get => _drawing != null ? _drawing.Top : 0;
        set
        {
            if (_drawing != null)
                _drawing.Top = value;
        }
    }

    public double Width
    {
        get => _drawing != null ? _drawing.Width : 0;
        set
        {
            if (_drawing != null)
                _drawing.Width = value;
        }
    }

    public double Height
    {
        get => _drawing != null ? _drawing.Height : 0;
        set
        {
            if (_drawing != null)
                _drawing.Height = value;
        }
    }


    public bool Visible
    {
        get => _drawing != null ? _drawing.Visible : false;
        set
        {
            if (_drawing != null)
                _drawing.Visible = value;
        }
    }

    public bool PrintObject
    {
        get => _drawing != null ? _drawing.PrintObject : false;
        set
        {
            if (_drawing != null)
                _drawing.PrintObject = value;
        }
    }

    public bool Locked
    {
        get => _drawing != null ? _drawing.Locked : false;
        set
        {
            if (_drawing != null)
                _drawing.Locked = value;
        }
    }

    public IExcelBorder? Border => _drawing != null ? new ExcelBorder(_drawing.Border) : null;

    public IExcelInterior? Interior => _drawing != null ? new ExcelInterior(_drawing.Interior) : null;


    public IExcelDrawingObjects? ParentDrawing => _drawing != null && _drawing.Parent is MsExcel.DrawingObjects dObj && dObj != null ? new ExcelDrawingObjects(dObj) : null;

    public IExcelWorksheet? Worksheet => _drawing != null && _drawing.Parent is MsExcel.Worksheet worksheet && worksheet != null ? new ExcelWorksheet(worksheet) : null;

    public IExcelRange? TopLeftCell => _drawing != null ? new ExcelRange(_drawing.TopLeftCell) : null;


    public string Text
    {
        get => _drawing != null ? _drawing.Text : string.Empty;
        set
        {
            if (_drawing != null)
                _drawing.Text = value;
        }
    }

    public IExcelFont? Font => _drawing != null ? new ExcelFont(_drawing.Font) : null;

    public XlHAlign HorizontalAlignment
    {
        get => _drawing != null ? _drawing.HorizontalAlignment.ObjectConvertEnum(XlHAlign.xlHAlignGeneral) : XlHAlign.xlHAlignGeneral;
        set
        {
            if (_drawing == null)
                return;
            _drawing.HorizontalAlignment = value.EnumConvert(MsExcel.XlHAlign.xlHAlignGeneral);
        }
    }

    public XlVAlign VerticalAlignment
    {
        get => _drawing != null ? _drawing.VerticalAlignment.ObjectConvertEnum(XlVAlign.xlVAlignBottom) : XlVAlign.xlVAlignBottom;
        set
        {
            if (_drawing == null)
                return;
            _drawing.VerticalAlignment = value.EnumConvert(MsExcel.XlVAlign.xlVAlignBottom);
        }

    }

    public IExcelShapeRange? ShapeRange
    {
        get => _drawing != null ? new ExcelShapeRange(_drawing.ShapeRange) : null;
    }

    internal ExcelDrawing(MsExcel.Drawing shape)
    {
        _drawing = shape ?? throw new ArgumentNullException(nameof(shape));
        _disposedValue = false;
    }

    public void Select(bool replace = true)
    {
        try
        {
            _drawing?.Select(replace);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法选择绘图对象。", ex);
        }
    }

    public void Delete()
    {
        try
        {
            _drawing?.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法删除绘图对象。", ex);
        }
    }

    public void Copy()
    {
        try
        {
            _drawing?.Copy();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法复制绘图对象。", ex);
        }
    }

    public void Cut()
    {
        try
        {
            _drawing?.Cut();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法剪切绘图对象。", ex);
        }
    }

    public void Move(double left, double top)
    {
        if (_drawing == null) return;
        try
        {
            _drawing.Left = left;
            _drawing.Top = top;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法移动绘图对象。", ex);
        }
    }

    public void Resize(double width, double height)
    {
        if (_drawing == null) return;
        try
        {
            _drawing.Width = width;
            _drawing.Height = height;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法调整绘图对象大小。", ex);
        }
    }


    /// <summary>
    /// 将图表对象置于最前面
    /// </summary>
    public void BringToFront()
    {
        _drawing?.BringToFront();
    }

    /// <summary>
    /// 将图表对象置于最后面
    /// </summary>
    public void SendToBack()
    {
        _drawing?.SendToBack();
    }


    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _drawing != null)
        {
            Marshal.ReleaseComObject(_drawing);
            _drawing = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

}