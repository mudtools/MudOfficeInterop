namespace MudTools.OfficeInterop.Word.Imps;


/// <summary>
/// Word.ChartColorFormat 的封装实现类。
/// </summary>
internal class WordChartColorFormat : IWordChartColorFormat
{
    private MsWord.ChartColorFormat _chartColorFormat;
    private bool _disposedValue;

    internal WordChartColorFormat(MsWord.ChartColorFormat chartColorFormat)
    {
        _chartColorFormat = chartColorFormat ?? throw new ArgumentNullException(nameof(chartColorFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _chartColorFormat != null ? new WordApplication(_chartColorFormat.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _chartColorFormat?.Parent;

    /// <inheritdoc/>
    public int RGB
    {
        get => _chartColorFormat?.RGB ?? 0;
    }

    /// <inheritdoc/>
    public int Type
    {
        get => _chartColorFormat?.Type ?? 0;
    }
    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _chartColorFormat != null)
        {
            Marshal.ReleaseComObject(_chartColorFormat);
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