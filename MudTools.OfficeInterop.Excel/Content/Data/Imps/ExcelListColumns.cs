//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

// =============================================
// 内部实现类：ExcelListColumns
// =============================================
internal class ExcelListColumns : IExcelListColumns
{
    internal MsExcel.ListColumns? _listColumns;
    private DisposableList _disposables = new();
    private bool _disposedValue = false;

    /// <summary>
    /// 构造函数，初始化封装对象。
    /// </summary>
    /// <param name="listColumns">原始的 COM ListColumns 对象</param>
    internal ExcelListColumns(MsExcel.ListColumns listColumns)
    {
        _listColumns = listColumns ?? throw new ArgumentNullException(nameof(listColumns));
    }

    /// <summary>
    /// 获取集合中列的总数。
    /// </summary>
    public int Count => _listColumns?.Count ?? 0;

    /// <summary>
    /// 通过索引（从 1 开始）获取指定的列。
    /// </summary>
    /// <param name="index">列索引（1-based）</param>
    /// <returns>对应的列对象</returns>
    public IExcelListColumn? this[int index]
    {
        get
        {
            if (_listColumns == null) return null;
            var col = new ExcelListColumn(_listColumns[index]);
            _disposables.Add(col);
            return col;
        }
    }

    /// <summary>
    /// 获取此集合所属的父对象（通常是 ListObject）。
    /// </summary>
    public object? Parent => _listColumns?.Parent;

    /// <summary>
    /// 获取此集合所属的 Excel 应用程序对象。
    /// </summary>
    public IExcelApplication? Application => _listColumns != null ? new ExcelApplication(_listColumns.Application) : null;

    /// <summary>
    /// 向集合中添加一个新列（插入在末尾或指定位置）。
    /// </summary>
    /// <param name="position">插入位置（可选，默认为末尾）</param>
    /// <returns>新创建的列对象</returns>
    public IExcelListColumn? Add(int? position = null)
    {
        if (_listColumns == null) return null;
        MsExcel.ListColumn newColumn;
        if (position.HasValue)
            newColumn = _listColumns.Add(position.Value);
        else
            newColumn = _listColumns.Add();

        var col = new ExcelListColumn(newColumn);
        _disposables.Add(col);
        return col;
    }

    #region IEnumerable<IExcelListColumn> Support

    /// <summary>
    /// 返回枚举器，用于 foreach 遍历。
    /// </summary>
    /// <returns>枚举器</returns>
    public IEnumerator<IExcelListColumn> GetEnumerator()
    {
        if (_listColumns == null)
            yield break;

        for (int i = 1; i <= _listColumns.Count; i++)
        {
            yield return this[i];
        }
    }

    /// <summary>
    /// 非泛型枚举器支持。
    /// </summary>
    /// <returns>枚举器</returns>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    #region IDisposable Support

    /// <summary>
    /// 释放托管和非托管资源。
    /// </summary>
    /// <param name="disposing">是否正在释放托管资源</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _listColumns != null)
        {
            Marshal.ReleaseComObject(_listColumns);
            _listColumns = null;
            _disposables.Dispose();
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 析构函数，确保即使未调用 Dispose 也能释放 COM 对象。
    /// </summary>
    ~ExcelListColumns()
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