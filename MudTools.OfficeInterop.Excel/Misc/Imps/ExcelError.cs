namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelError : IExcelError
{
    private MsExcel.Error _error;
    private bool _disposedValue;

    public object Parent => _error.Parent;

    public IExcelApplication Application => new ExcelApplication(_error.Application);

    public object Value => _error.Value;

    public bool Ignore
    {
        get => _error.Ignore;
        set => _error.Ignore = value;
    }

    internal ExcelError(MsExcel.Error error)
    {
        _error = error ?? throw new ArgumentNullException(nameof(error));
        _disposedValue = false;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _error != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_error) > 0) { }
            }
            catch { }
            _error = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}