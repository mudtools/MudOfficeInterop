//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Trendline 的封装实现类。
/// </summary>
internal class WordChartTrendline : IWordChartTrendline
{
    private MsWord.Trendline _trendline;
    private bool _disposedValue;

    internal WordChartTrendline(MsWord.Trendline trendline)
    {
        _trendline = trendline ?? throw new ArgumentNullException(nameof(trendline));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _trendline != null ? new WordApplication(_trendline.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _trendline?.Parent;

    /// <inheritdoc/>
    public string Name => _trendline?.Name ?? string.Empty;

    /// <inheritdoc/>
    public int Index => _trendline?.Index ?? 0;

    /// <inheritdoc/>
    public XlTrendlineType Type
    {
        get => _trendline?.Type != null ? (XlTrendlineType)(int)_trendline?.Type : XlTrendlineType.xlLinear;
        set
        {
            if (_trendline != null) _trendline.Type = (MsWord.XlTrendlineType)(int)value;
        }
    }

    /// <inheritdoc/>
    public int Order
    {
        get => _trendline?.Order ?? 2;
        set
        {
            if (_trendline != null)
                _trendline.Order = value;
        }
    }

    /// <inheritdoc/>
    public int Period
    {
        get => _trendline?.Period ?? 2;
        set
        {
            if (_trendline != null)
                _trendline.Period = value;
        }
    }

    /// <inheritdoc/>
    public bool DisplayEquation
    {
        get => _trendline?.DisplayEquation ?? false;
        set
        {
            if (_trendline != null)
                _trendline.DisplayEquation = value;
        }
    }

    /// <inheritdoc/>
    public bool DisplayRSquared
    {
        get => _trendline?.DisplayRSquared ?? false;
        set
        {
            if (_trendline != null)
                _trendline.DisplayRSquared = value;
        }
    }

    /// <inheritdoc/>
    public double Forward
    {
        get => _trendline?.Forward ?? 0.0;
        set
        {
            if (_trendline != null)
                _trendline.Forward = value;
        }
    }

    /// <inheritdoc/>
    public double Backward
    {
        get => _trendline?.Backward ?? 0.0;
        set
        {
            if (_trendline != null)
                _trendline.Backward = value;
        }
    }

    /// <inheritdoc/>
    public double Intercept
    {
        get => _trendline?.Intercept ?? 0.0;
        set
        {
            if (_trendline != null)
                _trendline.Intercept = value;
        }
    }
    #endregion

    #region 对象属性实现
    /// <inheritdoc/>
    public IWordChartBorder? Border => _trendline?.Border != null ? new WordChartBorder(_trendline.Border) : null;

    /// <inheritdoc/>
    public IWordChartFormat? Format => _trendline?.Format != null ? new WordChartFormat(_trendline.Format) : null;

    /// <inheritdoc/>
    public IWordChartDataLabel? DataLabel => _trendline?.DataLabel != null ? new WordChartDataLabel(_trendline.DataLabel) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _trendline?.Select();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _trendline?.Delete();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_trendline != null)
            {
                Marshal.ReleaseComObject(_trendline);
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