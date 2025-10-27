//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Series 对象的二次封装实现类
/// 实现 IExcelSeries 接口
/// </summary>
internal class ExcelSeries : IExcelSeries
{
    internal MsExcel.Series? _series;
    private bool _disposedValue = false;

    internal ExcelSeries(MsExcel.Series series)
    {
        _series = series ?? throw new ArgumentNullException(nameof(series));
    }

    #region 基础属性
    public string Name
    {
        get => _series?.Name ?? string.Empty;
        set
        {
            if (_series == null)
                return;
            _series.Name = value;
        }
    }
    public object? Parent => _series?.Parent;

    public IExcelApplication? Application => _series != null ? new ExcelApplication(_series.Application) : null;

    public MsoChartType ChartType
    {
        get => _series?.ChartType.EnumConvert(MsoChartType.xl3DColumn) ?? MsoChartType.xl3DColumn;
        set
        {
            if (_series == null)
                return;
            _series.ChartType = value.EnumConvert(MsExcel.XlChartType.xl3DColumn);
        }
    }

    public XlAxisGroup AxisGroup
    {
        get => _series?.AxisGroup.EnumConvert(XlAxisGroup.xlPrimary) ?? XlAxisGroup.xlPrimary;
        set
        {
            if (_series == null)
                return;
            _series.AxisGroup = value.EnumConvert(MsExcel.XlAxisGroup.xlPrimary);
        }
    }

    public string Formula
    {
        get => _series?.Formula ?? string.Empty;
        set
        {
            if (_series != null)
                _series.Formula = value;
        }
    }

    public string FormulaLocal
    {
        get => _series?.FormulaLocal ?? string.Empty;
        set
        {
            if (_series != null)
                _series.FormulaLocal = value;
        }
    }

    public string FormulaR1C1
    {
        get => _series?.FormulaR1C1 ?? string.Empty;
        set
        {
            if (_series != null)
                _series.FormulaR1C1 = value;
        }
    }

    public string FormulaR1C1Local
    {
        get => _series?.FormulaR1C1Local ?? string.Empty;
        set
        {
            if (_series != null)
                _series.FormulaR1C1Local = value;
        }
    }
    #endregion

    #region 数据属性
    public object? XValues
    {
        get => _series?.XValues ?? null;
        set
        {
            if (_series != null)
                _series.XValues = value;// value 可以是 object[] 或 MsExcel.Range
        }
    }

    public object? Values
    {
        get => _series?.Values ?? null;
        set
        {
            if (_series != null)
                _series.Values = value;// value 可以是 object[] 或 MsExcel.Range
        }
    }

    public object? BubbleSizes
    {
        get => _series?.BubbleSizes ?? null;
        set
        {
            if (_series != null)
                _series.BubbleSizes = value;// value 可以是 object[] 或 MsExcel.Range
        }
    }
    #endregion

    #region 格式设置
    public IExcelChartFormat? Format => _series != null ? new ExcelChartFormat(_series.Format) : null;

    public IExcelChartFillFormat? Fill => _series != null ? new ExcelChartFillFormat(_series.Fill) : null;

    public IExcelBorder? Border => _series != null ? new ExcelBorder(_series.Border) : null;

    public XlMarkerStyle MarkerStyle
    {
        get => _series?.MarkerStyle.EnumConvert(XlMarkerStyle.xlMarkerStyleAutomatic) ?? XlMarkerStyle.xlMarkerStyleAutomatic;
        set
        {
            if (_series == null)
                return;
            _series.MarkerStyle = value.EnumConvert(MsExcel.XlMarkerStyle.xlMarkerStyleAutomatic);
        }
    }

    public int MarkerSize
    {
        get => _series?.MarkerSize ?? 0;
        set
        {
            if (_series != null)
                _series.MarkerSize = value;
        }
    }

    public int MarkerBackgroundColor
    {
        get => _series?.MarkerBackgroundColor ?? 0;
        set
        {
            if (_series != null)
                _series.MarkerBackgroundColor = value;
        }
    }

    public XlColorIndex MarkerBackgroundColorIndex
    {
        get => _series?.MarkerBackgroundColorIndex.EnumConvert(XlColorIndex.xlColorIndexAutomatic) ?? XlColorIndex.xlColorIndexAutomatic;
        set
        {
            if (_series == null)
                return;
            _series.MarkerBackgroundColorIndex = value.EnumConvert(MsExcel.XlColorIndex.xlColorIndexAutomatic);
        }
    }


    public int MarkerForegroundColor
    {
        get => _series?.MarkerForegroundColor ?? 0;
        set
        {
            if (_series != null)
                _series.MarkerForegroundColor = value;
        }
    }

    public XlColorIndex MarkerForegroundColorIndex
    {
        get => _series?.MarkerForegroundColorIndex.EnumConvert(XlColorIndex.xlColorIndexAutomatic) ?? XlColorIndex.xlColorIndexAutomatic;
        set
        {
            if (_series == null)
                return;
            _series.MarkerForegroundColorIndex = value.EnumConvert(MsExcel.XlColorIndex.xlColorIndexAutomatic);
        }
    }

    public bool Smooth
    {
        get => _series?.Smooth ?? false;
        set
        {
            if (_series != null)
                _series.Smooth = value;
        }
    }

    public int PlotOrder
    {
        get => _series?.PlotOrder ?? 0;
        set
        {
            if (_series != null)
                _series.PlotOrder = value;
        }
    }
    #endregion

    #region 状态属性
    public bool HasLeaderLines
    {
        get => _series?.HasLeaderLines ?? false;
        set
        {
            if (_series != null)
                _series.HasLeaderLines = value;
        }
    }

    public bool HasDataLabels
    {
        get => _series?.HasDataLabels ?? false;
        set
        {
            if (_series != null)
                _series.HasDataLabels = value;
        }
    }

    public bool HasErrorBars
    {
        get => _series?.HasErrorBars ?? false;
        set
        {
            if (_series != null)
                _series.HasErrorBars = value;
        }
    }

    public bool IsProtected => false;
    #endregion

    #region 图表元素 (子对象)

    /// <summary>
    /// 获取样式的内部格式对象
    /// </summary>
    public IExcelInterior? Interior => _series != null ? new ExcelInterior(_series.Interior) : null;

    public IExcelErrorBars? ErrorBars => _series != null ? new ExcelErrorBars(_series.ErrorBars) : null;

    #endregion

    #region 操作方法

    public IExcelErrorBars? ErrorBar(
                            XlErrorBarDirection direction, XlErrorBarInclude include,
                            XlErrorBarType type,
                            object? amount = null,
                            object? minusValues = null)
    {
        if (_series == null)
            return null;

        var amountObj = Type.Missing;
        var minusValuesObj = Type.Missing;
        if (amount != null && amount is IExcelRange range)
        {
            amountObj = range.Value;
        }
        else if (amount != null)
        {
            amountObj = amount;
        }
        if (minusValues != null && minusValues is IExcelRange mrange)
        {
            minusValuesObj = mrange.Value;
        }
        else if (minusValues != null)
        {
            minusValuesObj = minusValues;
        }

        var errorObj = _series.ErrorBar(direction.EnumConvert(MsExcel.XlErrorBarDirection.xlY),
                             include.EnumConvert(MsExcel.XlErrorBarInclude.xlErrorBarIncludeBoth),
                             type.EnumConvert(MsExcel.XlErrorBarType.xlErrorBarTypeCustom),
                             amountObj, minusValuesObj);
        if (errorObj is MsExcel.ErrorBars errorBar)
            return new ExcelErrorBars(errorBar);
        return null;
    }

    public IExcelTrendlines? Trendlines()
    {
        if (_series == null)
            return null;
        var trends = _series.Trendlines();
        if (trends != null && trends is MsExcel.Trendlines trendlines)
            return new ExcelTrendlines(trendlines);
        return null;
    }

    public IExcelTrendline? Trendlines(XlTrendlineType trendlineType)
    {
        if (_series == null)
            return null;
        var trends = _series.Trendlines(trendlineType.EnumConvert(MsExcel.XlTrendlineType.xlExponential));
        if (trends != null && trends is MsExcel.Trendline trendline)
        {
            return new ExcelTrendline(trendline);
        }
        return null;
    }

    public IExcelDataLabels? DataLabels()
    {
        if (_series == null)
            return null;
        if (!_series.HasDataLabels)
            return null;

        var labels = _series.DataLabels();
        if (labels != null && labels is MsExcel.DataLabels dataLabels)
            return new ExcelDataLabels(dataLabels);
        return null;
    }

    public IExcelDataLabel? DataLabels(object obj)
    {
        if (_series == null)
            return null;
        if (!_series.HasDataLabels)
            return null;
        var label = _series.DataLabels(obj);
        if (label != null && label is MsExcel.DataLabel dataLabel)
            return new ExcelDataLabel(dataLabel);
        return null;
    }

    public void Select()
    {
        if (_series == null)
            return;
        _series.Select();
    }

    public void Delete()
    {
        if (_series == null)
            return;
        _series.Delete();
    }

    public void ClearFormats()
    {
        if (_series == null)
            return;
        _series.ClearFormats();
    }

    public void Copy()
    {
        if (_series == null)
            return;
        _series.Copy();
    }
    #endregion

    #region 图表操作
    public void ApplyDataLabels(XlDataLabelsType type = XlDataLabelsType.xlDataLabelsShowValue,
                                bool legendKey = false, bool autoText = true,
                                bool hasLeaderLines = false, bool showSeriesName = false,
                                bool showCategoryName = false, bool showValue = true,
                                bool showPercentage = false, bool showBubbleSize = false,
                                string? separator = null)
    {
        if (_series == null)
            return;
        _series.ApplyDataLabels(
            (MsExcel.XlDataLabelsType)type,
            legendKey,
            autoText,
            hasLeaderLines,
            showSeriesName,
            showCategoryName,
            showValue,
            showPercentage,
            showBubbleSize,
            separator.ComArgsVal()
        );
    }

    #endregion

    #region 格式设置方法   

    public void SetMarker(XlMarkerStyle style, int size, int backgroundColor, int foregroundColor)
    {
        if (_series == null)
            return;
        MarkerStyle = style;
        MarkerSize = size;
        MarkerBackgroundColor = backgroundColor;
        MarkerForegroundColor = foregroundColor;
    }
    #endregion


    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_series != null)
                Marshal.ReleaseComObject(_series);
            _series = null;
        }
        _disposedValue = true;
    }

    ~ExcelSeries()
    {
        Dispose(disposing: false);
    }

    public void Dispose()
    {

        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
    #endregion
}
