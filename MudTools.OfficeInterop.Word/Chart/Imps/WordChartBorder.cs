namespace MudTools.OfficeInterop.Word.Imps;


/// <summary>
/// Word.ChartBorder 的封装实现类。
/// </summary>
internal class WordChartBorder : IWordChartBorder
{
    private MsWord.ChartBorder _chartBorder;
    private bool _disposedValue;

    internal WordChartBorder(MsWord.ChartBorder chartBorder)
    {
        _chartBorder = chartBorder ?? throw new ArgumentNullException(nameof(chartBorder));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _chartBorder != null ? new WordApplication(_chartBorder.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _chartBorder?.Parent;

    /// <inheritdoc/>
    public object Color
    {
        get => _chartBorder?.Color;
        set
        {
            if (_chartBorder != null)
                _chartBorder.Color = value;
        }
    }

    /// <inheritdoc/>
    public XlColorIndex ColorIndex
    {
        get => _chartBorder?.ColorIndex != null ? (XlColorIndex)(int)_chartBorder?.ColorIndex : XlColorIndex.xlColorIndexNone;
        set
        {
            if (_chartBorder != null) _chartBorder.ColorIndex = (MsCore.XlColorIndex)(int)value;
        }
    }

    /// <inheritdoc/>
    public XlLineStyle LineStyle
    {
        get => _chartBorder?.LineStyle != null ? (XlLineStyle)(int)_chartBorder?.LineStyle : XlLineStyle.xlLineStyleNone;
        set
        {
            if (_chartBorder != null) _chartBorder.LineStyle = (MsWord.XlLineStyle)(int)value;
        }
    }

    /// <inheritdoc/>
    public float Weight
    {
        get => _chartBorder?.Weight.ConvertToFloat() ?? 0f;
        set
        {
            if (_chartBorder != null)
                _chartBorder.Weight = value;
        }
    }
    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _chartBorder != null)
        {
            Marshal.ReleaseComObject(_chartBorder);
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