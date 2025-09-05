namespace MudTools.OfficeInterop.Word.Imps;


/// <summary>
/// Word.ChartFormat 的封装实现类。
/// </summary>
internal class WordChartFormat : IWordChartFormat
{
    private MsWord.ChartFormat _chartFormat;
    private bool _disposedValue;

    internal WordChartFormat(MsWord.ChartFormat chartFormat)
    {
        _chartFormat = chartFormat ?? throw new ArgumentNullException(nameof(chartFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _chartFormat != null ? new WordApplication(_chartFormat.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _chartFormat?.Parent;

    /// <inheritdoc/>
    public IWordFillFormat? Fill => _chartFormat?.Fill != null ? new WordFillFormat(_chartFormat.Fill) : null;

    /// <inheritdoc/>
    public IWordLineFormat? Line => _chartFormat?.Line != null ? new WordLineFormat(_chartFormat.Line) : null;

    /// <inheritdoc/>
    public IWordShadowFormat? Shadow => _chartFormat?.Shadow != null ? new WordShadowFormat(_chartFormat.Shadow) : null;

    /// <inheritdoc/>
    public IWordGlowFormat? Glow => _chartFormat?.Glow != null ? new WordGlowFormat(_chartFormat.Glow) : null;

    /// <inheritdoc/>
    public IWordThreeDFormat? ThreeD => _chartFormat?.ThreeD != null ? new WordThreeDFormat(_chartFormat.ThreeD) : null;

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_chartFormat != null)
            {
                Marshal.ReleaseComObject(_chartFormat);
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