namespace MudTools.OfficeInterop.Word.Imps;


/// <summary>
/// Word.ChartFont 的封装实现类。
/// </summary>
internal class WordChartFont : IWordChartFont
{
    private MsWord.ChartFont _chartFont;
    private bool _disposedValue;

    internal WordChartFont(MsWord.ChartFont chartFont)
    {
        _chartFont = chartFont ?? throw new ArgumentNullException(nameof(chartFont));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _chartFont != null ? new WordApplication(_chartFont.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _chartFont?.Parent;

    /// <inheritdoc/>
    public string Name
    {
        get => _chartFont?.Name.ToString() ?? string.Empty;
        set
        {
            if (_chartFont != null)
                _chartFont.Name = value;
        }
    }

    /// <inheritdoc/>
    public float Size
    {
        get => _chartFont?.Size.ConvertToFloat() ?? 0f;
        set
        {
            if (_chartFont != null)
                _chartFont.Size = value;
        }
    }

    /// <inheritdoc/>
    public bool Bold
    {
        get => _chartFont?.Bold.ConvertToBool() ?? false;
        set
        {
            if (_chartFont != null)
                _chartFont.Bold = value;
        }
    }

    /// <inheritdoc/>
    public bool Italic
    {
        get => _chartFont?.Italic.ConvertToBool() ?? false;
        set
        {
            if (_chartFont != null)
                _chartFont.Italic = value;
        }
    }

    /// <inheritdoc/>
    public bool Underline
    {
        get => _chartFont?.Underline.ConvertToBool() ?? false;
        set
        {
            if (_chartFont != null)
                _chartFont.Underline = value;
        }
    }

    /// <inheritdoc/>
    public object Color
    {
        get => _chartFont?.Color;
        set
        {
            if (_chartFont != null)
                _chartFont.Color = value;
        }
    }

    /// <inheritdoc/>
    public XlColorIndex ColorIndex
    {
        get => _chartFont?.ColorIndex != null ? (XlColorIndex)(int)_chartFont?.ColorIndex : XlColorIndex.xlColorIndexNone;
        set
        {
            if (_chartFont != null) _chartFont.ColorIndex = (MsCore.XlColorIndex)(int)value;
        }
    }
    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _chartFont != null)
        {
            Marshal.ReleaseComObject(_chartFont);
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