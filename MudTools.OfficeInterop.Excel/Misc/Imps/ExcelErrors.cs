namespace MudTools.OfficeInterop.Excel.Imps;


internal class ExcelErrors : IExcelErrors
{
    private MsExcel.Errors _errors;
    private bool _disposedValue;

    public object Parent => _errors.Parent;

    public IExcelApplication Application => new ExcelApplication(_errors.Application);


    public IExcelError this[object index] => new ExcelError(_errors[index]);

    internal ExcelErrors(MsExcel.Errors errors)
    {
        _errors = errors ?? throw new ArgumentNullException(nameof(errors));
        _disposedValue = false;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _errors != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_errors) > 0) { }
            }
            catch { }
            _errors = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}