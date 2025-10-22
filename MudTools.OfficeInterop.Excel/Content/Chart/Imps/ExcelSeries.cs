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
    internal MsExcel.Series _series;
    private bool _disposedValue = false;

    internal ExcelSeries(MsExcel.Series series)
    {
        _series = series ?? throw new ArgumentNullException(nameof(series));
    }

    #region 基础属性
    public string Name
    {
        get => _series.Name;
        set => _series.Name = value;
    }
    public object Parent => _series.Parent;

    public IExcelApplication Application => new ExcelApplication(_series.Application);

    public int ChartType
    {
        get => (int)_series.ChartType;
        set => _series.ChartType = (MsExcel.XlChartType)value;
    }

    public string Formula
    {
        get => _series.Formula;
        set => _series.Formula = value;
    }

    public string FormulaLocal
    {
        get => _series.FormulaLocal;
        set => _series.FormulaLocal = value;
    }

    public string FormulaR1C1
    {
        get => _series.FormulaR1C1;
        set => _series.FormulaR1C1 = value;
    }

    public string FormulaR1C1Local
    {
        get => _series.FormulaR1C1Local;
        set => _series.FormulaR1C1Local = value;
    }
    #endregion

    #region 数据属性
    public object XValues
    {
        get => _series.XValues;
        set => _series.XValues = value; // value 可以是 object[] 或 MsExcel.Range
    }

    public object Values
    {
        get => _series.Values;
        set => _series.Values = value; // value 可以是 object[] 或 MsExcel.Range
    }

    public object BubbleSizes
    {
        get => _series.BubbleSizes;
        set => _series.BubbleSizes = value; // value 可以是 object[] 或 MsExcel.Range
    }
    #endregion

    #region 格式设置
    public IExcelChartFormat Format => new ExcelChartFormat(_series.Format);

    public IExcelChartFillFormat Fill => new ExcelChartFillFormat(_series.Fill);

    public IExcelBorder Border => new ExcelBorder(_series.Border);

    public int MarkerStyle
    {
        get => (int)_series.MarkerStyle;
        set => _series.MarkerStyle = (MsExcel.XlMarkerStyle)value;
    }

    public int MarkerSize
    {
        get => _series.MarkerSize;
        set => _series.MarkerSize = value;
    }

    public int MarkerBackgroundColor
    {
        get => _series.MarkerBackgroundColor;
        set => _series.MarkerBackgroundColor = value;
    }

    public XlColorIndex MarkerBackgroundColorIndex
    {
        get => (XlColorIndex)_series.MarkerBackgroundColorIndex;
        set => _series.MarkerBackgroundColorIndex = (MsExcel.XlColorIndex)value;
    }

    public int MarkerForegroundColor
    {
        get => _series.MarkerForegroundColor;
        set => _series.MarkerForegroundColor = value;
    }

    public XlColorIndex MarkerForegroundColorIndex
    {
        get => (XlColorIndex)_series.MarkerForegroundColorIndex;
        set => _series.MarkerForegroundColorIndex = (MsExcel.XlColorIndex)value;
    }

    public bool Smooth
    {
        get => _series.Smooth;
        set => _series.Smooth = value;
    }

    public int PlotOrder
    {
        get => _series.PlotOrder;
        set => _series.PlotOrder = value;
    }
    #endregion

    #region 状态属性
    public bool HasLeaderLines
    {
        get => _series.HasLeaderLines;
        set => _series.HasLeaderLines = value;
    }

    public void DataLabels()
    {
        _series.DataLabels();
    }

    public bool HasDataLabels
    {
        get => _series.HasDataLabels;
        set => _series.HasDataLabels = value;
    }

    public bool HasErrorBars
    {
        get => _series.HasErrorBars;
        set => _series.HasErrorBars = value;
    }

    public bool IsProtected => false;
    #endregion

    #region 图表元素 (子对象)

    /// <summary>
    /// 获取样式的内部格式对象
    /// </summary>
    public IExcelInterior Interior => new ExcelInterior(_series.Interior);

    public IExcelTrendlines Trendlines => new ExcelTrendlines((MsExcel.Trendlines)_series.Trendlines());

    public IExcelErrorBars ErrorBars => new ExcelErrorBars(_series.ErrorBars);

    #endregion

    #region 操作方法
    public void Select()
    {
        _series.Select();
    }

    public void Delete()
    {
        _series.Delete();
    }

    public void ClearFormats()
    {
        _series.ClearFormats();
    }

    public void Copy()
    {
        _series.Copy();
    }
    #endregion

    #region 图表操作
    public void ApplyDataLabels(int type = 2, bool legendKey = false, bool autoText = true,
                                bool hasLeaderLines = false, bool showSeriesName = false,
                                bool showCategoryName = false, bool showValue = true,
                                bool showPercentage = false, bool showBubbleSize = false,
                                object separator = null)
    {
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
            separator
        );
    }

    #endregion

    #region 格式设置方法   

    public void SetMarker(int style, int size, int backgroundColor, int foregroundColor)
    {
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
            try
            {
                // 释放底层COM对象
                if (_series != null)
                    Marshal.ReleaseComObject(_series);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
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
