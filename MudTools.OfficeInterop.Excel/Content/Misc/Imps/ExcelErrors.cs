namespace MudTools.OfficeInterop.Excel.Imps;


internal class ExcelErrors : IExcelErrors
{
    private MsExcel.Errors? _errors;
    private bool _disposedValue;
    private DisposableList _disposables = [];

    public object? Parent => _errors?.Parent;

    public IExcelApplication? Application => new ExcelApplication(_errors.Application);


    public IExcelError? this[object index]
    {
        get
        {
            if (_errors == null)
                return null;
            var error = _errors[index];
            if (error == null)
                return null;
            var excelError = new ExcelError(error);
            _disposables.Add(excelError);
            return excelError;
        }
    }

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
            Marshal.ReleaseComObject(_errors)
            _disposables.Dispose();
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