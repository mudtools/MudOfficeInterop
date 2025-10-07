
namespace MudTools.OfficeInterop.Excel.Imps;


/// <summary>
/// UpBars COM 对象的封装实现类。
/// </summary>
internal class ExcelUpBars : IExcelUpBars
{
    internal MsExcel.UpBars? _upBars;
    private bool _disposedValue;

    internal ExcelUpBars(MsExcel.UpBars upBars)
    {
        _upBars = upBars ?? throw new ArgumentNullException(nameof(upBars));
        _disposedValue = false;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        if (disposing)
        {
            if (_upBars != null)
            {
                Marshal.ReleaseComObject(_upBars);
                _upBars = null;
            }
        }
        _disposedValue = true;
    }

    public void Dispose() => Dispose(true);

    public object? Parent => _upBars?.Parent;
    public IExcelApplication? Application =>
        _upBars?.Application != null ? new ExcelApplication(_upBars.Application as MsExcel.Application) : null;


    public IExcelBorder? Border =>
        _upBars?.Border != null ? new ExcelBorder(_upBars.Border) : null;

    public IExcelChartFillFormat? Fill =>
        _upBars?.Fill != null ? new ExcelChartFillFormat(_upBars.Fill) : null;

    public IExcelInterior? Interior =>
         _upBars?.Interior != null ? new ExcelInterior(_upBars.Interior) : null;

    public IExcelChartFormat? Format =>
        _upBars?.Format != null ? new ExcelChartFormat(_upBars.Format) : null;

    public void Select() => _upBars?.Select();
    public void Delete() => _upBars?.Delete();
}