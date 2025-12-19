//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Legend 的封装实现类。
/// </summary>
internal class WordLegend : IWordLegend
{
    private MsWord.Legend _legend;
    private bool _disposedValue;

    internal WordLegend(MsWord.Legend legend)
    {
        _legend = legend ?? throw new ArgumentNullException(nameof(legend));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _legend != null ? new WordApplication(_legend.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _legend?.Parent;

    /// <inheritdoc/>
    public string Name
    {
        get => _legend?.Name ?? string.Empty;
    }


    /// <inheritdoc/>
    public MsWord.XlLegendPosition Position
    {
        get => _legend?.Position ?? MsWord.XlLegendPosition.xlLegendPositionRight;
        set
        {
            if (_legend != null)
                _legend.Position = value;
        }
    }

    /// <inheritdoc/>
    public bool IncludeInLayout
    {
        get => _legend?.IncludeInLayout ?? true;
        set
        {
            if (_legend != null)
                _legend.IncludeInLayout = value;
        }
    }

    /// <inheritdoc/>
    public double Left
    {
        get => _legend?.Left ?? 0.0;
        set
        {
            if (_legend != null)
                _legend.Left = value;
        }
    }

    /// <inheritdoc/>
    public double Top
    {
        get => _legend?.Top ?? 0.0;
        set
        {
            if (_legend != null)
                _legend.Top = value;
        }
    }

    /// <inheritdoc/>
    public double Width
    {
        get => _legend?.Width ?? 0.0;
        set
        {
            if (_legend != null)
                _legend.Width = value;
        }
    }

    /// <inheritdoc/>
    public double Height
    {
        get => _legend?.Height ?? 0.0;
        set
        {
            if (_legend != null)
                _legend.Height = value;
        }
    }

    /// <inheritdoc/>
    public bool AutoScaleFont
    {
        get => _legend.AutoScaleFont.ConvertToBool();
        set
        {
            if (_legend != null)
                _legend.AutoScaleFont = value;
        }
    }

    /// <inheritdoc/>
    public bool Shadow
    {
        get => _legend?.Shadow ?? false;
        set
        {
            if (_legend != null)
                _legend.Shadow = value;
        }
    }

    #endregion

    #region 对象属性实现 
    /// <inheritdoc/>
    public IWordChartFont Font => _legend?.Font != null ? new WordChartFont(_legend.Font) : null;

    /// <inheritdoc/>
    public IWordInterior Interior => _legend?.Interior != null ? new WordInterior(_legend.Interior) : null;

    /// <inheritdoc/>
    public IWordChartFillFormat Fill => _legend?.Fill != null ? new WordChartFillFormat(_legend.Fill) : null;

    /// <inheritdoc/>
    public IWordChartBorder Border => _legend?.Border != null ? new WordChartBorder(_legend.Border) : null;

    /// <inheritdoc/>
    public IWordChartFormat Format => _legend?.Format != null ? new WordChartFormat(_legend.Format) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _legend?.Select();
    }

    /// <inheritdoc/>
    public void Clear()
    {
        _legend?.Clear();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _legend?.Delete();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_legend != null)
            {
                Marshal.ReleaseComObject(_legend);
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