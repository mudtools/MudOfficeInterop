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
    internal MsExcel.ListColumn _listColumn;
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
    public object Parent => _listColumn.Parent;

    /// <summary>
    /// 获取此列所属的 Excel 应用程序对象。
    /// </summary>
    public IExcelApplication Application => new ExcelApplication(_listColumn.Application);

    /// <summary>
    /// 获取此列在 ListColumns 集合中的索引（从 1 开始）。
    /// </summary>
    public int Index => _listColumn.Index;

    /// <summary>
    /// 获取或设置此列的标题名称（即表头）。
    /// </summary>
    public string Name
    {
        get => _listColumn.Name;
        set => _listColumn.Name = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>
    /// 获取此列对应的数据范围（DataBodyRange），不包含标题。
    /// 如果表无数据行，则可能为 null。
    /// </summary>
    public IExcelRange? DataBodyRange
    {
        get
        {
            var range = _listColumn.DataBodyRange;
            return range != null ? new ExcelRange(range) : null;
        }
    }

    /// <summary>
    /// 获取此列的整个范围（包括标题和数据体）。
    /// </summary>
    public IExcelRange? Range
    {
        get
        {
            var range = _listColumn.Range;
            return range != null ? new ExcelRange(range) : null;
        }
    }

    /// <summary>
    /// 获取或设置此列的总计行公式（仅当 ListObject.ShowTotals = true 时有效）。
    /// </summary>
    public XlTotalsCalculation TotalsCalculation
    {
        get => _listColumn.TotalsCalculation.EnumConvert(XlTotalsCalculation.xlTotalsCalculationCount);
        set => _listColumn.TotalsCalculation = value.EnumConvert(MsExcel.XlTotalsCalculation.xlTotalsCalculationCount);
    }

    public IExcelListDataFormat? ListDataFormat
    {
        get
        {
            var format = _listColumn.ListDataFormat;
            return format != null ? new ExcelListDataFormat(format) : null;
        }

    }

    /// <summary>
    /// 删除此列（将从表格中移除该列）。
    /// </summary>
    public void Delete()
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelListColumn));
        try
        {
            _listColumn.Delete();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"删除列 '{this.Name}' 失败: {ex.Message}");
            throw; // 重新抛出，让调用方决定如何处理
        }
    }

    /// <summary>
    /// 获取此列对应的总计行单元格（仅当启用总计行时有效）。
    /// </summary>
    public IExcelRange Total
    {
        get
        {
            var range = _listColumn.Total;
            return range != null ? new ExcelRange(range) : null;
        }
    }

    #region IDisposable Support

    /// <summary>
    /// 释放托管和非托管资源。
    /// </summary>
    /// <param name="disposing">是否正在释放托管资源</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_listColumn != null)
            {
                Marshal.ReleaseComObject(_listColumn);
                _listColumn = null;
            }
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