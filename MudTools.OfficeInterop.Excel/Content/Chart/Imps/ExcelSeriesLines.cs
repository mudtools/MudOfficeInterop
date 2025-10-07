
namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// SeriesLines COM 对象的封装实现类。
/// 负责管理 COM 对象生命周期，提供安全的属性访问和资源释放。
/// </summary>
internal class ExcelSeriesLines : IExcelSeriesLines
{
    /// <summary>
    /// 内部持有的原始 COM 对象。
    /// </summary>
    internal MsExcel.SeriesLines? _seriesLines;

    /// <summary>
    /// 标记对象是否已被释放。
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装类。
    /// </summary>
    /// <param name="seriesLines">原始的 SeriesLines COM 对象，不可为 null。</param>
    /// <exception cref="ArgumentNullException">当传入的 seriesLines 为 null 时抛出。</exception>
    internal ExcelSeriesLines(MsExcel.SeriesLines seriesLines)
    {
        _seriesLines = seriesLines ?? throw new ArgumentNullException(nameof(seriesLines));
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的受保护虚方法，支持派生类重写。
    /// </summary>
    /// <param name="disposing">是否由用户代码显式调用释放。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放托管资源：释放 COM 对象
            if (_seriesLines != null)
            {
                Marshal.ReleaseComObject(_seriesLines);
                _seriesLines = null;
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 公开的 Dispose 方法，用于显式释放资源。
    /// 调用后对象不应再被使用。
    /// </summary>
    public void Dispose() => Dispose(true);

    /// <summary>
    /// 获取此对象的父对象（通常是 Chart）。
    /// </summary>
    public object? Parent => _seriesLines?.Parent;

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// 返回封装后的 <see cref="IExcelApplication"/> 接口实例。
    /// </summary>
    public IExcelApplication? Application =>
        _seriesLines?.Application != null
            ? new ExcelApplication(_seriesLines.Application as MsExcel.Application)
            : null;


    /// <summary>
    /// 获取系列连线的边框格式（用于设置颜色、线型、粗细等）。
    /// </summary>
    public IExcelBorder? Border =>
        _seriesLines != null
            ? new ExcelBorder(_seriesLines.Border)
            : null;

    public IExcelChartFormat? Format =>
        _seriesLines != null
            ? new ExcelChartFormat(_seriesLines.Format)
            : null;

    /// <summary>
    /// 选中此系列连线（激活并高亮显示）。
    /// </summary>
    public void Select()
    {
        _seriesLines?.Select();
    }

    /// <summary>
    /// 删除此系列连线（将其设为不可见，并从图表中移除）。
    /// </summary>
    public void Delete()
    {
        _seriesLines?.Delete();
    }
}