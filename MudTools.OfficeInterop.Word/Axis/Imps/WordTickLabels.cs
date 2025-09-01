//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.TickLabels 的封装实现类。
/// </summary>
internal class WordTickLabels : IWordTickLabels
{
    private MsWord.TickLabels _tickLabels;
    private bool _disposedValue;

    internal WordTickLabels(MsWord.TickLabels tickLabels)
    {
        _tickLabels = tickLabels ?? throw new ArgumentNullException(nameof(tickLabels));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _tickLabels != null ? new WordApplication(_tickLabels.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _tickLabels?.Parent;

    /// <inheritdoc/>
    public string Name => _tickLabels?.Name ?? string.Empty;


    /// <inheritdoc/>
    public XlTickLabelOrientation Orientation
    {
        get => _tickLabels?.Orientation != null ? (XlTickLabelOrientation)(int)_tickLabels?.Orientation : XlTickLabelOrientation.xlTickLabelOrientationAutomatic;
        set
        {
            if (_tickLabels != null) _tickLabels.Orientation = (MsWord.XlTickLabelOrientation)(int)value;
        }
    }

    /// <inheritdoc/>
    public bool AutoScaleFont
    {
        get => _tickLabels.AutoScaleFont.ConvertToBool();
        set
        {
            if (_tickLabels != null)
                _tickLabels.AutoScaleFont = value;
        }
    }

    /// <inheritdoc/>
    public string NumberFormat
    {
        get => _tickLabels?.NumberFormat ?? string.Empty;
        set
        {
            if (_tickLabels != null)
                _tickLabels.NumberFormat = value;
        }
    }

    /// <inheritdoc/>
    public bool NumberFormatLinked
    {
        get => _tickLabels?.NumberFormatLinked ?? true;
        set
        {
            if (_tickLabels != null)
                _tickLabels.NumberFormatLinked = value;
        }
    }

    /// <inheritdoc/>
    public object NumberFormatLocal
    {
        get => _tickLabels?.NumberFormatLocal;
        set
        {
            if (_tickLabels != null)
                _tickLabels.NumberFormatLocal = value;
        }
    }

    /// <inheritdoc/>
    public int ReadingOrder
    {
        get => _tickLabels?.ReadingOrder ?? 0;
        set
        {
            if (_tickLabels != null)
                _tickLabels.ReadingOrder = value;
        }
    }

    /// <inheritdoc/>
    public int Depth
    {
        get => _tickLabels?.Depth ?? 0;
    }

    /// <inheritdoc/>
    public int Offset
    {
        get => _tickLabels?.Offset ?? 0;
        set
        {
            if (_tickLabels != null)
                _tickLabels.Offset = value;
        }
    }

    /// <inheritdoc/>
    public int Alignment
    {
        get => _tickLabels?.Alignment ?? 0;
        set
        {
            if (_tickLabels != null)
                _tickLabels.Alignment = value;
        }
    }

    /// <inheritdoc/>
    public bool MultiLevel
    {
        get => _tickLabels?.MultiLevel ?? true;
        set
        {
            if (_tickLabels != null)
                _tickLabels.MultiLevel = value;
        }
    }
    #endregion

    #region 对象属性实现
    /// <inheritdoc/>
    public IWordChartFont Font => _tickLabels?.Font != null ? new WordChartFont(_tickLabels.Font) : null;

    /// <inheritdoc/>
    public IWordChartFormat Format => _tickLabels?.Format != null ? new WordChartFormat(_tickLabels.Format) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _tickLabels?.Select();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _tickLabels?.Delete();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_tickLabels != null)
            {
                Marshal.ReleaseComObject(_tickLabels);
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