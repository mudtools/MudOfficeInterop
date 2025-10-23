//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

// =============================================
// 内部实现类：ExcelListColumn
// =============================================
internal class ExcelListColumn : IExcelListColumn
{
    internal MsExcel.ListColumn? _listColumn;
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelListColumn));
    private bool _disposedValue = false;

    /// <summary>
    /// 构造函数，初始化封装对象。
    /// </summary>
    /// <param name="listColumn">原始的 COM ListColumn 对象</param>
    internal ExcelListColumn(MsExcel.ListColumn listColumn)
    {
        _listColumn = listColumn ?? throw new ArgumentNullException(nameof(listColumn));
    }

    /// <summary>
    /// 获取此列所属的父对象（通常是 ListObject）。
    /// </summary>
    public object? Parent => _listColumn?.Parent;

    /// <summary>
    /// 获取此列所属的 Excel 应用程序对象。
    /// </summary>
    public IExcelApplication? Application => _listColumn != null ? new ExcelApplication(_listColumn.Application) : null;

    /// <summary>
    /// 获取此列在 ListColumns 集合中的索引（从 1 开始）。
    /// </summary>
    public int Index => _listColumn?.Index ?? 0;

    /// <summary>
    /// 获取或设置此列的标题名称（即表头）。
    /// </summary>
    public string Name
    {
        get => _listColumn?.Name ?? string.Empty;
        set
        {
            if (_listColumn != null)
                _listColumn.Name = value;
        }
    }

    /// <summary>
    /// 获取此列对应的数据范围（DataBodyRange），不包含标题。
    /// 如果表无数据行，则可能为 null。
    /// </summary>
    public IExcelRange? DataBodyRange => _listColumn != null ? new ExcelRange(_listColumn.DataBodyRange) : null;

    /// <summary>
    /// 获取此列的整个范围（包括标题和数据体）。
    /// </summary>
    public IExcelRange? Range => _listColumn != null ? new ExcelRange(_listColumn.Range) : null;

    /// <summary>
    /// 获取或设置此列的总计行公式（仅当 ListObject.ShowTotals = true 时有效）。
    /// </summary>
    public XlTotalsCalculation TotalsCalculation
    {
        get => _listColumn != null ? _listColumn.TotalsCalculation.EnumConvert(XlTotalsCalculation.xlTotalsCalculationCount) : XlTotalsCalculation.xlTotalsCalculationCount;
        set
        {
            if (_listColumn != null)
                _listColumn.TotalsCalculation = value.EnumConvert(MsExcel.XlTotalsCalculation.xlTotalsCalculationCount);
        }
    }

    public IExcelListDataFormat? ListDataFormat => _listColumn != null ? new ExcelListDataFormat(_listColumn.ListDataFormat) : null;

    /// <summary>
    /// 删除此列（将从表格中移除该列）。
    /// </summary>
    public void Delete()
    {
        if (_listColumn == null) throw new ObjectDisposedException(nameof(ExcelListColumn));
        try
        {
            _listColumn.Delete();
        }
        catch (Exception ex)
        {
            log.Error($"删除列 '{this.Name}' 失败: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// 获取此列对应的总计行单元格（仅当启用总计行时有效）。
    /// </summary>
    public IExcelRange? Total => _listColumn != null ? new ExcelRange(_listColumn.Total) : null;

    #region IDisposable Support

    /// <summary>
    /// 释放托管和非托管资源。
    /// </summary>
    /// <param name="disposing">是否正在释放托管资源</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _listColumn != null)
        {
            Marshal.ReleaseComObject(_listColumn);
            _listColumn = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 析构函数，确保即使未调用 Dispose 也能释放 COM 对象。
    /// </summary>
    ~ExcelListColumn()
    {
        Dispose(disposing: false);
    }

    /// <summary>
    /// 公开的 Dispose 方法，用于显式释放资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    #endregion
}