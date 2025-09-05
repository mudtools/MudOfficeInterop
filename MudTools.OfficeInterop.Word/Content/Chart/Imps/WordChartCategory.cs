namespace MudTools.OfficeInterop.Word.Imps;


/// <summary>
/// Word.ChartCategory 的封装实现类。
/// </summary>
internal class WordChartCategory : IWordChartCategory
{
    private MsWord.ChartCategory _chartCategory;
    private bool _disposedValue;

    internal WordChartCategory(MsWord.ChartCategory chartCategory)
    {
        _chartCategory = chartCategory ?? throw new ArgumentNullException(nameof(chartCategory));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public object Parent => _chartCategory?.Parent;

    /// <inheritdoc/>
    public string Name => _chartCategory?.Name ?? string.Empty;

    /// <inheritdoc/>
    public bool IsFiltered
    {
        get => _chartCategory?.IsFiltered != null && _chartCategory.IsFiltered;
        set
        {
            if (_chartCategory != null)
                _chartCategory.IsFiltered = value;
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _chartCategory != null)
        {
            Marshal.ReleaseComObject(_chartCategory);
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