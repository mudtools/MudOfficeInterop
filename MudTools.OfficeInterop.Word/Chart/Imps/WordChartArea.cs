//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.ChartArea 的封装实现类。
/// </summary>
internal class WordChartArea : IWordChartArea
{
    private MsWord.ChartArea _chartArea;
    private bool _disposedValue;

    internal WordChartArea(MsWord.ChartArea chartArea)
    {
        _chartArea = chartArea ?? throw new ArgumentNullException(nameof(chartArea));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _chartArea != null ? new WordApplication(_chartArea.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _chartArea?.Parent;

    /// <inheritdoc/>
    public string Name
    {
        get => _chartArea?.Name ?? string.Empty;
    }

    /// <inheritdoc/>
    public bool AutoScaleFont
    {
        get => _chartArea.AutoScaleFont.ConvertToBool();
        set
        {
            if (_chartArea != null)
                _chartArea.AutoScaleFont = value;
        }
    }

    /// <inheritdoc/>
    public double Left
    {
        get => _chartArea?.Left ?? 0.0;
        set
        {
            if (_chartArea != null)
                _chartArea.Left = value;
        }
    }

    /// <inheritdoc/>
    public double Top
    {
        get => _chartArea?.Top ?? 0.0;
        set
        {
            if (_chartArea != null)
                _chartArea.Top = value;
        }
    }

    /// <inheritdoc/>
    public double Width
    {
        get => _chartArea?.Width ?? 0.0;
        set
        {
            if (_chartArea != null)
                _chartArea.Width = value;
        }
    }

    /// <inheritdoc/>
    public double Height
    {
        get => _chartArea?.Height ?? 0.0;
        set
        {
            if (_chartArea != null)
                _chartArea.Height = value;
        }
    }

    /// <inheritdoc/>
    public object InteriorColor
    {
        get => _chartArea?.Interior.Color;
        set
        {
            if (_chartArea?.Interior != null)
                _chartArea.Interior.Color = value;
        }
    }

    /// <inheritdoc/>
    public object BorderColor
    {
        get => _chartArea?.Border.Color;
        set
        {
            if (_chartArea?.Border != null)
                _chartArea.Border.Color = value;
        }
    }

    /// <inheritdoc/>
    public bool Fill
    {
        get => _chartArea?.Fill.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_chartArea?.Fill != null)
                _chartArea.Fill.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public bool Shadow
    {
        get => _chartArea?.Shadow ?? false;
        set
        {
            if (_chartArea != null)
                _chartArea.Shadow = value;
        }
    }
    #endregion

    #region 对象属性实现
    /// <inheritdoc/>
    public IWordChartFont? Font => _chartArea?.Font != null ? new WordChartFont(_chartArea.Font) : null;

    /// <inheritdoc/>
    public IWordInterior? Interior => _chartArea?.Interior != null ? new WordInterior(_chartArea.Interior) : null;

    /// <inheritdoc/>
    public IWordChartFillFormat? FillFormat => _chartArea?.Fill != null ? new WordChartFillFormat(_chartArea.Fill) : null;

    /// <inheritdoc/>
    public IWordChartBorder? Border => _chartArea?.Border != null ? new WordChartBorder(_chartArea.Border) : null;

    /// <inheritdoc/>
    public IWordChartFormat? Format => _chartArea?.Format != null ? new WordChartFormat(_chartArea.Format) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _chartArea?.Select();
    }

    /// <inheritdoc/>
    public void Clear()
    {
        _chartArea?.Clear();
    }

    /// <inheritdoc/>
    public void ClearContents()
    {
        _chartArea?.ClearContents();
    }

    /// <inheritdoc/>
    public void ClearFormats()
    {
        _chartArea?.ClearFormats();
    }

    /// <inheritdoc/>
    public void Copy()
    {
        _chartArea?.Copy();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_chartArea != null)
            {
                Marshal.ReleaseComObject(_chartArea);
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