
namespace MudTools.OfficeInterop.Excel.Imps;


/// <summary>
/// HiLoLines COM 对象的封装实现类。
/// </summary>
internal class ExcelHiLoLines : IExcelHiLoLines
{
    internal MsExcel.HiLoLines _hiLoLines;
    private bool _disposedValue;

    internal ExcelHiLoLines(MsExcel.HiLoLines hiLoLines)
    {
        _hiLoLines = hiLoLines ?? throw new ArgumentNullException(nameof(hiLoLines));
        _disposedValue = false;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        if (disposing)
        {
            if (_hiLoLines != null)
            {
                Marshal.ReleaseComObject(_hiLoLines);
                _hiLoLines = null;
            }
        }
        _disposedValue = true;
    }

    public void Dispose() => Dispose(true);

    public object? Parent => _hiLoLines?.Parent;
    public IExcelApplication? Application =>
        _hiLoLines?.Application != null ? new ExcelApplication(_hiLoLines.Application as MsExcel.Application) : null;



    public IExcelBorder? Border =>
        _hiLoLines?.Border != null ? new ExcelBorder(_hiLoLines.Border) : null;

    public IExcelChartFormat? Format =>
    _hiLoLines?.Format != null ? new ExcelChartFormat(_hiLoLines.Format) : null;

    public void Select() => _hiLoLines?.Select();
    public void Delete() => _hiLoLines?.Delete();
}