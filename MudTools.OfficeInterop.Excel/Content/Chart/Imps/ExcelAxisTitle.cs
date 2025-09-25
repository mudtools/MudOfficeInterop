
namespace MudTools.OfficeInterop.Excel.Imps;


/// <summary>
/// AxisTitle COM 对象的封装实现类。
/// 负责管理 COM 对象生命周期，提供安全的属性访问和资源释放。
/// </summary>
internal class ExcelAxisTitle : IExcelAxisTitle
{
    /// <summary>
    /// 内部持有的原始 COM 对象。
    /// </summary>
    internal MsExcel.AxisTitle _axisTitle;

    /// <summary>
    /// 标记对象是否已被释放。
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装类。
    /// </summary>
    /// <param name="axisTitle">原始的 AxisTitle COM 对象，不可为 null。</param>
    /// <exception cref="ArgumentNullException">当传入的 axisTitle 为 null 时抛出。</exception>
    internal ExcelAxisTitle(MsExcel.AxisTitle axisTitle)
    {
        _axisTitle = axisTitle ?? throw new ArgumentNullException(nameof(axisTitle));
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的受保护虚方法，支持派生类重写。
    /// </summary>
    /// <param name="disposing">是否由用户代码显式调用释放。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放托管资源：释放 COM 对象
            if (_axisTitle != null)
            {
                Marshal.ReleaseComObject(_axisTitle);
                _axisTitle = null;
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 公开的 Dispose 方法，用于显式释放资源。
    /// 调用后对象不应再被使用。
    /// </summary>
    public void Dispose() => Dispose(true);

    /// <summary>
    /// 获取此对象的父对象（通常是 Axis）。
    /// </summary>
    public object? Parent => _axisTitle?.Parent;

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// 返回封装后的 <see cref="IExcelApplication"/> 接口实例。
    /// </summary>
    public IExcelApplication? Application =>
        _axisTitle?.Application != null
            ? new ExcelApplication(_axisTitle.Application as MsExcel.Application)
            : null;

    public string Name => _axisTitle != null ? _axisTitle.Name : "";

    /// <summary>
    /// 获取或设置坐标轴标题的文本内容。
    /// </summary>
    public string Text
    {
        get => _axisTitle?.Text ?? string.Empty;
        set
        {
            if (_axisTitle != null && value != null)
                _axisTitle.Text = value;
        }
    }

    public string Caption
    {
        get => _axisTitle?.Caption ?? string.Empty;
        set
        {
            if (_axisTitle != null && value != null)
                _axisTitle.Caption = value;
        }
    }

    public double Left
    {
        get => _axisTitle?.Left ?? 0;
        set
        {
            if (_axisTitle != null)
                _axisTitle.Left = value;
        }
    }

    public double Top
    {
        get => _axisTitle?.Top ?? 0;
        set
        {
            if (_axisTitle != null)
                _axisTitle.Top = value;
        }
    }

    public double Width
    {
        get => _axisTitle?.Width ?? 0;
    }

    public double Height
    {
        get => _axisTitle?.Height ?? 0;
    }

    public bool Shadow
    {
        get => _axisTitle?.Shadow ?? false;
        set
        {
            if (_axisTitle != null)
                _axisTitle.Shadow = value;
        }
    }

    public int ReadingOrder
    {
        get => _axisTitle?.ReadingOrder ?? 0;
        set
        {
            if (_axisTitle != null)
                _axisTitle.ReadingOrder = value;
        }
    }

    public bool IncludeInLayout
    {
        get => _axisTitle?.IncludeInLayout ?? false;
        set
        {
            if (_axisTitle != null)
                _axisTitle.IncludeInLayout = value;
        }
    }

    public string Formula
    {
        get => _axisTitle?.Formula ?? "";
        set
        {
            if (_axisTitle != null)
                _axisTitle.Formula = value;
        }
    }

    public string FormulaR1C1
    {
        get => _axisTitle?.FormulaR1C1 ?? "";
        set
        {
            if (_axisTitle != null)
                _axisTitle.FormulaR1C1 = value;
        }
    }

    public string FormulaLocal
    {
        get => _axisTitle?.FormulaLocal ?? "";
        set
        {
            if (_axisTitle != null)
                _axisTitle.FormulaLocal = value;
        }
    }

    public string FormulaR1C1Local
    {
        get => _axisTitle?.FormulaR1C1Local ?? "";
        set
        {
            if (_axisTitle != null)
                _axisTitle.FormulaR1C1Local = value;
        }
    }

    public XlChartElementPosition Position
    {
        get => _axisTitle != null ? _axisTitle.Position.EnumConvert(XlChartElementPosition.xlChartElementPositionAutomatic) : XlChartElementPosition.xlChartElementPositionAutomatic;
        set
        {
            if (_axisTitle != null)
                _axisTitle.Position = value.EnumConvert(MsExcel.XlChartElementPosition.xlChartElementPositionAutomatic);
        }
    }


    /// <summary>
    /// 获取坐标轴标题的文本方向（角度或预设方向）。
    /// 默认值：xlHorizontal
    /// </summary>
    public XlOrientation Orientation
    {
        get => _axisTitle != null
            ? _axisTitle.Orientation.ObjectConvertEnum(XlOrientation.xlHorizontal)
            : XlOrientation.xlHorizontal;

        set
        {
            if (_axisTitle != null)
                _axisTitle.Orientation = value;
        }
    }

    /// <summary>
    /// 获取或设置坐标轴标题的水平对齐方式。
    /// 默认值：xlHAlignCenter
    /// </summary>
    public XlHAlign HorizontalAlignment
    {
        get => _axisTitle != null
            ? _axisTitle.HorizontalAlignment.ObjectConvertEnum(XlHAlign.xlHAlignCenter)
            : XlHAlign.xlHAlignCenter;

        set
        {
            if (_axisTitle != null)
                _axisTitle.HorizontalAlignment = value;
        }
    }

    /// <summary>
    /// 获取或设置坐标轴标题的垂直对齐方式。
    /// 默认值：xlVAlignCenter
    /// </summary>
    public XlVAlign VerticalAlignment
    {
        get => _axisTitle != null
            ? _axisTitle.VerticalAlignment.ObjectConvertEnum(XlVAlign.xlVAlignCenter)
            : XlVAlign.xlVAlignCenter;

        set
        {
            if (_axisTitle != null)
                _axisTitle.VerticalAlignment = value;
        }
    }

    /// <summary>
    /// 获取坐标轴标题的字体格式。
    /// </summary>
    public IExcelFont? Font =>
        _axisTitle?.Font != null
            ? new ExcelFont(_axisTitle.Font)
            : null;

    public IExcelInterior? Interior
    {
        get
        {
            if (_axisTitle != null)
                return new ExcelInterior(_axisTitle.Interior);
            else
                return null;
        }
    }

    /// <summary>
    /// 获取坐标轴标题的字符格式（用于高级文本格式）。
    /// </summary>
    public IExcelCharacters? Characters =>
        _axisTitle?.Characters != null
            ? new ExcelCharacters(_axisTitle.Characters)
            : null;

    /// <summary>
    /// 获取坐标轴标题的填充格式。
    /// </summary>
    public IExcelChartFillFormat? Fill =>
        _axisTitle?.Fill != null
            ? new ExcelChartFillFormat(_axisTitle.Fill)
            : null;

    /// <summary>
    /// 获取坐标轴标题的边框格式。
    /// </summary>
    public IExcelBorder? Border =>
        _axisTitle != null
            ? new ExcelBorder(_axisTitle.Border)
            : null;

    /// <summary>
    /// 选中此坐标轴标题（激活并高亮显示）。
    /// </summary>
    public void Select()
    {
        _axisTitle?.Select();
    }

    /// <summary>
    /// 删除此坐标轴标题（将其设为不可见，并从图表中移除）。
    /// </summary>
    public void Delete()
    {
        _axisTitle?.Delete();
    }
}