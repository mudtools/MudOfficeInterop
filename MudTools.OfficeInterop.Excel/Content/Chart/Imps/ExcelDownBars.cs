
namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// DownBars COM 对象的封装实现类。
/// </summary>
internal class ExcelDownBars : IExcelDownBars
{
    internal MsExcel.DownBars? _downBars;
    private bool _disposedValue;

    internal ExcelDownBars(MsExcel.DownBars downBars)
    {
        _downBars = downBars ?? throw new ArgumentNullException(nameof(downBars));
        _disposedValue = false;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        if (disposing)
        {
            if (_downBars != null)
            {
                Marshal.ReleaseComObject(_downBars);
                _downBars = null;
            }
        }
        _disposedValue = true;
    }

    public void Dispose() => Dispose(true);

    public object? Parent => _downBars?.Parent;
    public IExcelApplication? Application =>
        _downBars?.Application != null ? new ExcelApplication(_downBars.Application as MsExcel.Application) : null;


    public IExcelBorder? Border =>
        _downBars?.Border != null ? new ExcelBorder(_downBars.Border) : null;

    public IExcelInterior? Interior
    => _downBars?.Interior != null ? new ExcelInterior(_downBars.Interior) : null;

    public IExcelChartFillFormat? Fill =>
        _downBars?.Fill != null ? new ExcelChartFillFormat(_downBars.Fill) : null;

    public IExcelChartFormat? Format =>
    _downBars?.Format != null ? new ExcelChartFormat(_downBars.Format) : null;

    public void Select() => _downBars?.Select();
    public void Delete() => _downBars?.Delete();
}
