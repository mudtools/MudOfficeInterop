//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
internal class ExcelDrawing : IExcelDrawing
{
    private MsExcel.Drawing _drawing;
    private bool _disposedValue;

    /// <summary>
    /// 获取绘图对象的索引
    /// </summary>
    public int Index => _drawing.Index;

    public string Name
    {
        get => _drawing.Name;
        set => _drawing.Name = value;
    }

    public double Left
    {
        get => _drawing.Left;
        set => _drawing.Left = (float)value;
    }

    public double Top
    {
        get => _drawing.Top;
        set => _drawing.Top = (float)value;
    }

    public double Width
    {
        get => _drawing.Width;
        set => _drawing.Width = (float)value;
    }

    public double Height
    {
        get => _drawing.Height;
        set => _drawing.Height = (float)value;
    }


    public bool Visible
    {
        get => _drawing.Visible;
        set => _drawing.Visible = value;
    }


    public bool Locked
    {
        get => _drawing.Locked;
        set => _drawing.Locked = !value;
    }


    public IExcelBorder Border
    {
        get => new ExcelBorder(_drawing.Border);
    }

    public IExcelInterior Interior
    {
        get => new ExcelInterior(_drawing.Interior);
    }


    public IExcelDrawingObjects Parent => new ExcelDrawingObjects(_drawing.Parent as MsExcel.DrawingObjects);

    public IExcelWorksheet Worksheet => new ExcelWorksheet(_drawing.Parent as MsExcel.Worksheet);


    public string Text
    {
        get => _drawing.Text;
        set => _drawing.Text = value;
    }

    public IExcelFont Font => new ExcelFont(_drawing.Font);

    public XlHAlign HorizontalAlignment
    {
        get => (XlHAlign)(int)_drawing.HorizontalAlignment;
        set => _drawing.HorizontalAlignment = (MsExcel.XlHAlign)value;
    }

    public XlVAlign VerticalAlignment
    {
        get => (XlVAlign)(int)_drawing.VerticalAlignment;
        set => _drawing.VerticalAlignment = (MsExcel.XlVAlign)value;
    }

    public IExcelShapeRange ShapeRange
    {
        get => new ExcelShapeRange(_drawing.ShapeRange);
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
            _drawing.Select(replace);
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
            _drawing.Delete();
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
            _drawing.Copy();
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
            _drawing.Cut();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法剪切绘图对象。", ex);
        }
    }

    public void Move(double left, double top)
    {
        try
        {
            _drawing.Left = (float)left;
            _drawing.Top = (float)top;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法移动绘图对象。", ex);
        }
    }

    public void Resize(double width, double height)
    {
        try
        {
            _drawing.Width = (float)width;
            _drawing.Height = (float)height;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法调整绘图对象大小。", ex);
        }
    }


    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _drawing != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_drawing) > 0) { }
            }
            catch { }
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