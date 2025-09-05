//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.ChartGroup 的封装实现类。
/// </summary>
internal class WordChartGroup : IWordChartGroup
{
    private MsWord.ChartGroup _chartGroup;
    private bool _disposedValue;

    internal WordChartGroup(MsWord.ChartGroup chartGroup)
    {
        _chartGroup = chartGroup ?? throw new ArgumentNullException(nameof(chartGroup));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _chartGroup != null ? new WordApplication(_chartGroup.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _chartGroup?.Parent;

    /// <inheritdoc/>
    public MsWord.XlAxisGroup AxisGroup
    {
        get => _chartGroup?.AxisGroup ?? MsWord.XlAxisGroup.xlPrimary;
        set
        {
            if (_chartGroup != null)
                _chartGroup.AxisGroup = value;
        }
    }

    /// <inheritdoc/>
    public bool HasSeriesLines
    {
        get => _chartGroup?.HasSeriesLines ?? false;
        set
        {
            if (_chartGroup != null)
                _chartGroup.HasSeriesLines = value;
        }
    }

    /// <inheritdoc/>
    public bool HasHiLoLines
    {
        get => _chartGroup?.HasHiLoLines ?? false;
        set
        {
            if (_chartGroup != null)
                _chartGroup.HasHiLoLines = value;
        }
    }

    /// <inheritdoc/>
    public bool HasUpDownBars
    {
        get => _chartGroup?.HasUpDownBars ?? false;
        set
        {
            if (_chartGroup != null)
                _chartGroup.HasUpDownBars = value;
        }
    }

    /// <inheritdoc/>
    public int Overlap
    {
        get => _chartGroup?.Overlap ?? 0;
        set
        {
            if (_chartGroup != null)
                _chartGroup.Overlap = value;
        }
    }

    /// <inheritdoc/>
    public int GapWidth
    {
        get => _chartGroup?.GapWidth ?? 150;
        set
        {
            if (_chartGroup != null)
                _chartGroup.GapWidth = value;
        }
    }


    /// <inheritdoc/>
    public bool HasRadarAxisLabels
    {
        get => _chartGroup?.HasRadarAxisLabels ?? false;
        set
        {
            if (_chartGroup != null)
                _chartGroup.HasRadarAxisLabels = value;
        }
    }


    /// <inheritdoc/>
    public bool Has3DShading
    {
        get => _chartGroup?.Has3DShading ?? false;
        set
        {
            if (_chartGroup != null)
                _chartGroup.Has3DShading = value;
        }
    }

    /// <inheritdoc/>
    public int BubbleScale
    {
        get => _chartGroup?.BubbleScale ?? 100;
        set
        {
            if (_chartGroup != null)
                _chartGroup.BubbleScale = value;
        }
    }

    /// <inheritdoc/>
    public bool ShowNegativeBubbles
    {
        get => _chartGroup?.ShowNegativeBubbles ?? false;
        set
        {
            if (_chartGroup != null)
                _chartGroup.ShowNegativeBubbles = value;
        }
    }

    /// <inheritdoc/>
    public XlSizeRepresents SizeRepresents
    {
        get => _chartGroup?.SizeRepresents != null ? (XlSizeRepresents)(int)_chartGroup?.SizeRepresents : XlSizeRepresents.xlSizeIsArea;
        set
        {
            if (_chartGroup != null) _chartGroup.SizeRepresents = (MsWord.XlSizeRepresents)(int)value;
        }
    }

    #endregion

    #region 对象属性实现

    /// <inheritdoc/>
    public IWordChartSeries? SeriesCollection => _chartGroup?.SeriesCollection() != null ?
        new WordChartSeries(_chartGroup.SeriesCollection() as MsWord.Series) : null;

    /// <inheritdoc/>
    public IWordChartHiLoLines? HiLoLines => _chartGroup?.HiLoLines != null ?
        new WordChartHiLoLines(_chartGroup.HiLoLines) : null;

    /// <inheritdoc/>
    public IWordTickLabels? RadarAxisLabels => _chartGroup?.RadarAxisLabels != null ?
        new WordTickLabels(_chartGroup.RadarAxisLabels) : null;

    /// <inheritdoc/>
    public IWordDownBars? DownBars => _chartGroup?.DownBars != null ?
        new WordDownBars(_chartGroup.DownBars) : null;

    /// <inheritdoc/>
    public IWordDropLines? DropLines => _chartGroup?.DropLines != null ?
        new WordDropLines(_chartGroup.DropLines) : null;

    /// <inheritdoc/>
    public IWordChartSeriesLines SeriesLines => _chartGroup?.SeriesLines != null ?
        new WordChartSeriesLines(_chartGroup.SeriesLines) : null;


    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_chartGroup != null)
            {
                Marshal.ReleaseComObject(_chartGroup);
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