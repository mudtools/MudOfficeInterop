
namespace MudTools.OfficeInterop.Excel.Imps;


/// <summary>
/// ChartGroup COM 对象的封装实现类。
/// 负责管理 COM 对象生命周期，提供安全的属性访问和资源释放。
/// </summary>
internal class ExcelChartGroup : IExcelChartGroup
{
    /// <summary>
    /// 内部持有的原始 COM 对象。
    /// </summary>
    internal MsExcel.ChartGroup? _chartGroup;

    /// <summary>
    /// 标记对象是否已被释放。
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装类。
    /// </summary>
    /// <param name="chartGroup">原始的 ChartGroup COM 对象，不可为 null。</param>
    /// <exception cref="ArgumentNullException">当传入的 chartGroup 为 null 时抛出。</exception>
    internal ExcelChartGroup(MsExcel.ChartGroup chartGroup)
    {
        _chartGroup = chartGroup ?? throw new ArgumentNullException(nameof(chartGroup));
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
            if (_chartGroup != null)
            {
                Marshal.ReleaseComObject(_chartGroup);
                _chartGroup = null;
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
    /// 获取此对象的父对象（通常是 Chart）。
    /// </summary>
    public object? Parent => _chartGroup?.Parent;

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// 返回封装后的 <see cref="IExcelApplication"/> 接口实例。
    /// </summary>
    public IExcelApplication? Application =>
        _chartGroup?.Application != null
            ? new ExcelApplication(_chartGroup.Application as MsExcel.Application)
            : null;

    /// <summary>
    /// 获取或设置系列之间的间隙宽度（0-500，100=默认）。
    /// </summary>
    public int GapWidth
    {
        get => _chartGroup?.GapWidth ?? 100;
        set
        {
            if (_chartGroup != null)
            {
                if (value < 0 || value > 500)
                    throw new ArgumentOutOfRangeException(nameof(value), "GapWidth 必须在 0 到 500 之间。");
                _chartGroup.GapWidth = value;
            }
        }
    }

    /// <summary>
    /// 获取或设置同组系列之间的重叠程度（-100 到 100）。
    /// </summary>
    public int Overlap
    {
        get => _chartGroup?.Overlap ?? 0;
        set
        {
            if (_chartGroup != null)
            {
                if (value < -100 || value > 100)
                    throw new ArgumentOutOfRangeException(nameof(value), "Overlap 必须在 -100 到 100 之间。");
                _chartGroup.Overlap = value;
            }
        }
    }

    /// <summary>
    /// 获取或设置是否显示高低点连线（Hi-Lo Lines）。
    /// </summary>
    public bool HasHiLoLines
    {
        get => _chartGroup != null && _chartGroup.HasHiLoLines;
        set
        {
            if (_chartGroup != null)
                _chartGroup.HasHiLoLines = value;
        }
    }

    /// <summary>
    /// 获取或设置是否显示垂线（Drop Lines）。
    /// </summary>
    public bool HasDropLines
    {
        get => _chartGroup != null && _chartGroup.HasDropLines;
        set
        {
            if (_chartGroup != null)
                _chartGroup.HasDropLines = value;
        }
    }

    /// <summary>
    /// 获取或设置是否显示上涨/下跌柱（Up/Down Bars）。
    /// </summary>
    public bool HasUpDownBars
    {
        get => _chartGroup != null && _chartGroup.HasUpDownBars;
        set
        {
            if (_chartGroup != null)
                _chartGroup.HasUpDownBars = value;
        }
    }

    public int SubType
    {
        get => _chartGroup?.SubType ?? 0;
        set
        {
            if (_chartGroup != null)
                _chartGroup.SubType = value;
        }
    }

    public int BubbleScale
    {
        get => _chartGroup?.BubbleScale ?? 0;
        set
        {
            if (_chartGroup != null)
                _chartGroup.BubbleScale = value;
        }
    }

    public bool ShowNegativeBubbles
    {
        get => _chartGroup?.ShowNegativeBubbles ?? false;
        set
        {
            if (_chartGroup != null)
                _chartGroup.ShowNegativeBubbles = value;
        }
    }

    public bool VaryByCategories
    {
        get => _chartGroup?.VaryByCategories ?? false;
        set
        {
            if (_chartGroup != null)
                _chartGroup.VaryByCategories = value;
        }
    }

    public XlSizeRepresents SizeRepresents
    {
        get => _chartGroup?.SizeRepresents.ObjectConvertEnum(XlSizeRepresents.xlSizeIsArea) ?? XlSizeRepresents.xlSizeIsArea;
        set
        {
            if (_chartGroup != null)
                _chartGroup.SizeRepresents = value.ObjectConvertEnum(MsExcel.XlSizeRepresents.xlSizeIsArea);
        }
    }

    public XlChartSplitType SplitType
    {
        get => _chartGroup?.SplitType.ObjectConvertEnum(XlChartSplitType.xlSplitByPosition) ?? XlChartSplitType.xlSplitByPosition;
        set
        {
            if (_chartGroup != null)
                _chartGroup.SplitType = value.ObjectConvertEnum(MsExcel.XlChartSplitType.xlSplitByPosition);
        }
    }

    public double SplitValue
    {
        get => _chartGroup?.SplitValue.ConvertToDouble() ?? 0;
        set
        {
            if (_chartGroup != null)
                _chartGroup.SplitValue = value;
        }
    }

    public int SecondPlotSize
    {
        get => _chartGroup?.SecondPlotSize ?? 0;
        set
        {
            if (_chartGroup != null)
                _chartGroup.SecondPlotSize = value;
        }
    }

    public bool Has3DShading
    {
        get => _chartGroup?.Has3DShading ?? false;
        set
        {
            if (_chartGroup != null)
                _chartGroup.Has3DShading = value;
        }
    }

    /// <summary>
    /// 获取图表组的类型（柱形图、折线图、饼图等）。
    /// </summary>
    public XlChartType Type =>
        _chartGroup != null
            ? _chartGroup.Type.ObjectConvertEnum(XlChartType.xlColumnClustered)
            : XlChartType.xlColumnClustered;

    /// <summary>
    /// 获取图表组的索引（从 1 开始）。
    /// </summary>
    public int Index => _chartGroup?.Index ?? 0;

    /// <summary>
    /// 获取图表组中的系列集合。
    /// </summary>
    public IExcelSeriesCollection? SeriesCollection(int? index = null)
    {
        if (_chartGroup != null)
        {
            var scObj = _chartGroup.SeriesCollection(index.ComArgsVal());
            if (scObj != null && scObj is MsExcel.SeriesCollection sc)
            {
                return new ExcelSeriesCollection(sc);
            }
        }
        return null;

    }

    public IExcelSeriesLines? SeriesLines =>
        _chartGroup?.SeriesLines != null
            ? new ExcelSeriesLines(_chartGroup.SeriesLines)
            : null;


    public IExcelTickLabels? RadarAxisLabels
    {
        get =>
            _chartGroup?.RadarAxisLabels != null
                ? new ExcelTickLabels(_chartGroup.RadarAxisLabels)
                : null;
    }

    /// <summary>
    /// 获取图表组的上涨柱（UpBars）格式。
    /// </summary>
    public IExcelUpBars? UpBars =>
        _chartGroup?.UpBars != null
            ? new ExcelUpBars(_chartGroup.UpBars)
            : null;

    /// <summary>
    /// 获取图表组的下跌柱（DownBars）格式。
    /// </summary>
    public IExcelDownBars? DownBars =>
        _chartGroup?.DownBars != null
            ? new ExcelDownBars(_chartGroup.DownBars)
            : null;

    /// <summary>
    /// 获取图表组的高低线（HiLoLines）格式。
    /// </summary>
    public IExcelHiLoLines? HiLoLines =>
        _chartGroup?.HiLoLines != null
            ? new ExcelHiLoLines(_chartGroup.HiLoLines)
            : null;

    /// <summary>
    /// 获取图表组的垂线（DropLines）格式。
    /// </summary>
    public IExcelDropLines? DropLines =>
        _chartGroup?.DropLines != null
            ? new ExcelDropLines(_chartGroup.DropLines)
            : null;
}