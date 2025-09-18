namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelAdjustments : IExcelAdjustments
{
    private MsExcel.Adjustments _adjustments;
    private bool _disposedValue;
    public object Parent => _adjustments.Parent;

    public int Count => _adjustments.Count;

    public ExcelAdjustments(MsExcel.Adjustments adjustments)
    {
        _adjustments = adjustments;
        _disposedValue = false;
    }

    public float this[int index]
    {
        get => _adjustments[index];
        set => _adjustments[index] = value;
    }

    public IEnumerator<float> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _adjustments != null)
        {
            Marshal.ReleaseComObject(_adjustments);
            _adjustments = null;
        }

        _disposedValue = true;
    }

    ~ExcelAdjustments()
    {
        Dispose(false);
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}