//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Chart 的封装实现类。
/// </summary>
internal class WordChart : IWordChart
{
    private MsWord.Chart _chart;
    private bool _disposedValue;

    internal WordChart(MsWord.Chart chart)
    {
        _chart = chart ?? throw new ArgumentNullException(nameof(chart));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _chart != null ? new WordApplication(_chart.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _chart?.Parent;

    /// <inheritdoc/>
    public MsoChartType ChartType
    {
        get => _chart?.ChartType != null ? (MsoChartType)(int)_chart?.ChartType : MsoChartType.xlLine;
        set
        {
            if (_chart != null) _chart.ChartType = (MsCore.XlChartType)(int)value;
        }
    }

    /// <inheritdoc/>
    public bool HasLegend
    {
        get => _chart?.HasLegend ?? false;
        set
        {
            if (_chart != null)
                _chart.HasLegend = value;
        }
    }

    /// <inheritdoc/>
    public bool HasDataTable
    {
        get => _chart?.HasDataTable ?? false;
        set
        {
            if (_chart != null)
                _chart.HasDataTable = value;
        }
    }

    /// <inheritdoc/>
    public bool HasTitle
    {
        get => _chart?.HasTitle ?? false;
        set
        {
            if (_chart != null)
                _chart.HasTitle = value;
        }
    }

    #endregion

    #region 对象属性实现

    /// <inheritdoc/>
    public IWordChartArea? ChartArea => _chart?.ChartArea != null ? new WordChartArea(_chart.ChartArea) : null;

    /// <inheritdoc/>
    public IWordPlotArea? PlotArea => _chart?.PlotArea != null ? new WordPlotArea(_chart.PlotArea) : null;

    /// <inheritdoc/>
    public IWordLegend? Legend => _chart?.Legend != null ? new WordLegend(_chart.Legend) : null;

    /// <inheritdoc/>
    public IWordChartTitle? ChartTitle => _chart?.ChartTitle != null ? new WordChartTitle(_chart.ChartTitle) : null;

    /// <inheritdoc/>
    public IWordDataTable? DataTable => _chart?.DataTable != null ? new WordDataTable(_chart.DataTable) : null;

    /// <inheritdoc/>
    public IWordChartData? ChartData => _chart?.ChartData != null ? new WordChartData(_chart.ChartData) : null;

    /// <inheritdoc/>
    public IWordChartSeriesCollection? SeriesCollection => _chart?.SeriesCollection() != null ?
        new WordChartSeriesCollection(_chart.SeriesCollection() as MsWord.SeriesCollection) : null;

    /// <inheritdoc/>
    public IWordChartGroups? ChartGroups => _chart?.ChartGroups != null ?
        new WordChartGroups(_chart.ChartGroups as MsWord.ChartGroups) : null;

    /// <inheritdoc/>
    public IWordAxis? CategoryAxis => _chart?.Axes(MsWord.XlAxisType.xlCategory) != null ?
        new WordAxis(_chart.Axes(MsWord.XlAxisType.xlCategory) as MsWord.Axis) : null;

    /// <inheritdoc/>
    public IWordAxis? ValueAxis => _chart?.Axes(MsWord.XlAxisType.xlValue) != null ?
        new WordAxis(_chart.Axes(MsWord.XlAxisType.xlValue) as MsWord.Axis) : null;

    /// <inheritdoc/>
    public IWordAxis? SecondaryCategoryAxis => _chart?.Axes(MsWord.XlAxisType.xlCategory, MsWord.XlAxisGroup.xlSecondary) != null ?
        new WordAxis(_chart.Axes(MsWord.XlAxisType.xlCategory, MsWord.XlAxisGroup.xlSecondary) as MsWord.Axis) : null;

    /// <inheritdoc/>
    public IWordAxis? SecondaryValueAxis => _chart?.Axes(MsWord.XlAxisType.xlValue, MsWord.XlAxisGroup.xlSecondary) != null ?
        new WordAxis(_chart.Axes(MsWord.XlAxisType.xlValue, MsWord.XlAxisGroup.xlSecondary) as MsWord.Axis) : null;

    /// <inheritdoc/>
    public IWordWalls? Walls => _chart?.Walls != null ? new WordWalls(_chart.Walls) : null;

    /// <inheritdoc/>
    public IWordFloor? Floor => _chart?.Floor != null ? new WordFloor(_chart.Floor) : null;
    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void ApplyChartStyle(MsoChartType style)
    {
        _chart?.ApplyCustomType((MsCore.XlChartType)(int)style);
    }

    /// <inheritdoc/>
    public void ApplyDataLabels(XlDataLabelsType type)
    {
        _chart?.ApplyDataLabels((MsWord.XlDataLabelsType)(int)type);
    }

    /// <inheritdoc/>
    public void SetSourceData(string source, XlRowCol plotBy)
    {
        _chart?.SetSourceData(source, (MsWord.XlRowCol)(int)plotBy);
    }

    /// <inheritdoc/>
    public void Select()
    {
        _chart?.Select();
    }

    /// <inheritdoc/>
    public void Copy()
    {
        _chart?.Copy();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _chart?.Delete();
    }

    /// <inheritdoc/>
    public void Refresh()
    {
        _chart?.Refresh();
    }

    /// <inheritdoc/>
    public void Export(string filename, string filterName, bool interactive)
    {
        _chart?.Export(filename, filterName, interactive);
    }

    /// <inheritdoc/>
    public void SetElement(MsoChartElementType element)
    {
        _chart?.SetElement((MsCore.MsoChartElementType)(int)element);
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_chart != null)
            {
                Marshal.ReleaseComObject(_chart);
            }
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}