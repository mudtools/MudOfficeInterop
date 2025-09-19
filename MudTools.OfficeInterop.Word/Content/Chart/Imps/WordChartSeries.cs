//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Series 的封装实现类。
/// </summary>
internal class WordChartSeries : IWordChartSeries
{
    private MsWord.Series _series;
    private bool _disposedValue;

    internal WordChartSeries(MsWord.Series series)
    {
        _series = series ?? throw new ArgumentNullException(nameof(series));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _series != null ? new WordApplication(_series.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _series?.Parent;

    /// <inheritdoc/>
    public string Name
    {
        get => _series?.Name ?? string.Empty;
        set
        {
            if (_series != null)
                _series.Name = value;
        }
    }

    /// <inheritdoc/>
    public object Values
    {
        get => _series?.Values;
        set
        {
            if (_series != null)
                _series.Values = value;
        }
    }

    /// <inheritdoc/>
    public object XValues
    {
        get => _series?.XValues;
        set
        {
            if (_series != null)
                _series.XValues = value;
        }
    }

    /// <inheritdoc/>
    public object BubbleSizes
    {
        get => _series?.BubbleSizes;
        set
        {
            if (_series != null)
                _series.BubbleSizes = value;
        }
    }

    /// <inheritdoc/>
    public MsoChartType ChartType
    {
        get => _series?.ChartType != null ? _series.ChartType.EnumConvert(MsoChartType.xlArea) : MsoChartType.xlArea;
        set
        {
            if (_series != null) _series.ChartType = value.EnumConvert(MsCore.XlChartType.xlArea);
        }
    }

    /// <inheritdoc/>
    public XlAxisGroup AxisGroup
    {
        get => _series?.AxisGroup != null ? _series.AxisGroup.EnumConvert(XlAxisGroup.xlPrimary) : XlAxisGroup.xlPrimary;
        set
        {
            if (_series != null) _series.AxisGroup = value.EnumConvert(MsWord.XlAxisGroup.xlPrimary);
        }

    }


    /// <inheritdoc/>
    public bool Smooth
    {
        get => _series?.Smooth ?? false;
        set
        {
            if (_series != null)
                _series.Smooth = value;
        }
    }

    /// <inheritdoc/>
    public XlMarkerStyle MarkerStyle
    {
        get => _series?.MarkerStyle != null ? _series.MarkerStyle.EnumConvert(XlMarkerStyle.xlMarkerStyleNone) : XlMarkerStyle.xlMarkerStyleNone;
        set
        {
            if (_series != null)
                _series.MarkerStyle = value.EnumConvert(MsWord.XlMarkerStyle.xlMarkerStyleNone);
        }
    }

    /// <inheritdoc/>
    public int MarkerSize
    {
        get => _series?.MarkerSize ?? 0;
        set
        {
            if (_series != null)
                _series.MarkerSize = value;
        }
    }

    /// <inheritdoc/>
    public int MarkerBackgroundColor
    {
        get => _series?.MarkerBackgroundColor != null ? _series.MarkerBackgroundColor : 0;
        set
        {
            if (_series != null)
                _series.MarkerBackgroundColor = value;
        }
    }

    /// <inheritdoc/>
    public int MarkerForegroundColor
    {
        get => _series?.MarkerForegroundColor != null ? _series.MarkerForegroundColor : 0;
        set
        {
            if (_series != null)
                _series.MarkerForegroundColor = value;
        }
    }

    /// <inheritdoc/>
    public bool HasDataLabels
    {
        get => _series?.HasDataLabels ?? false;
        set
        {
            if (_series != null)
                _series.HasDataLabels = value;
        }
    }

    /// <inheritdoc/>
    public bool HasErrorBars
    {
        get => _series?.HasErrorBars ?? false;
        set
        {
            if (_series != null)
                _series.HasErrorBars = value;
        }
    }

    #endregion

    #region 对象属性实现

    /// <inheritdoc/>
    public IWordChartFormat? Format => _series?.Format != null ? new WordChartFormat(_series.Format) : null;

    /// <inheritdoc/>
    public IWordChartDataLabels? DataLabels => _series?.DataLabels() != null ? new WordChartDataLabels(_series.DataLabels() as MsWord.DataLabels) : null;

    /// <inheritdoc/>
    public IWordChartTrendlines? Trendlines => _series?.Trendlines() != null ? new WordChartTrendlines(_series.Trendlines() as MsWord.Trendlines) : null;

    /// <inheritdoc/>
    public IWordErrorBars? ErrorBars => _series?.ErrorBars != null ? new WordErrorBars(_series.ErrorBars) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void ApplyDataLabels(MsWord.XlDataLabelsType type, bool legendKey, bool autoText, bool hasLeaderLines)
    {
        _series?.ApplyDataLabels(type, legendKey, autoText, hasLeaderLines);
    }

    /// <inheritdoc/>
    public void Select()
    {
        _series?.Select();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _series?.Delete();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放所有子对象
            if (_series != null)
            {
                Marshal.ReleaseComObject(_series);
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