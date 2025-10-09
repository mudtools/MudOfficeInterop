
namespace MudTools.OfficeInterop.Excel.Imps;


/// <summary>
/// DropLines COM 对象的封装实现类。
/// </summary>
internal class ExcelDropLines : IExcelDropLines
{
    internal MsExcel.DropLines _dropLines;
    private bool _disposedValue;

    internal ExcelDropLines(MsExcel.DropLines dropLines)
    {
        _dropLines = dropLines ?? throw new ArgumentNullException(nameof(dropLines));
        _disposedValue = false;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        if (disposing)
        {
            if (_dropLines != null)
            {
                Marshal.ReleaseComObject(_dropLines);
                _dropLines = null;
            }
        }
        _disposedValue = true;
    }

    public void Dispose() => Dispose(true);

    public object? Parent => _dropLines?.Parent;
    public IExcelApplication? Application =>
        _dropLines?.Application != null ? new ExcelApplication(_dropLines.Application as MsExcel.Application) : null;


    public IExcelBorder? Border =>
        _dropLines?.Border != null ? new ExcelBorder(_dropLines.Border) : null;

    public IExcelChartFormat? Format
    => _dropLines?.Format != null ? new ExcelChartFormat(_dropLines.Format) : null;

    public void Select() => _dropLines?.Select();
    public void Delete() => _dropLines?.Delete();
}