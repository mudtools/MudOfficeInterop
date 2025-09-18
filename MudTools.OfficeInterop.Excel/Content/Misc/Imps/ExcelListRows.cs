//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

// =============================================
// 内部实现类：ExcelListRows
// =============================================
internal class ExcelListRows : IExcelListRows
{
    internal MsExcel.ListRows _listRows;
    private bool _disposedValue = false;

    /// <summary>
    /// 构造函数，初始化封装对象。
    /// </summary>
    /// <param name="listRows">原始的 COM ListRows 对象</param>
    internal ExcelListRows(MsExcel.ListRows listRows)
    {
        _listRows = listRows ?? throw new ArgumentNullException(nameof(listRows));
    }

    /// <summary>
    /// 获取集合中行的总数。
    /// </summary>
    public int Count => _listRows.Count;

    /// <summary>
    /// 通过索引（从 1 开始）获取指定的行。
    /// </summary>
    /// <param name="index">行索引（1-based）</param>
    /// <returns>对应的行对象</returns>
    public IExcelListRow this[int index]
    {
        get
        {
            if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelListRows));
            return new ExcelListRow(_listRows[index]);
        }
    }

    /// <summary>
    /// 获取此集合所属的父对象（通常是 ListObject）。
    /// </summary>
    public object Parent => _listRows.Parent;

    /// <summary>
    /// 获取此集合所属的 Excel 应用程序对象。
    /// </summary>
    public IExcelApplication Application => new ExcelApplication(_listRows.Application);

    /// <summary>
    /// 向集合中添加一个新行（插入在末尾）。
    /// 新行将继承表格结构，所有单元格初始为空。
    /// </summary>
    /// <returns>新创建的行对象</returns>
    public IExcelListRow Add()
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelListRows));
        MsExcel.ListRow newRow = _listRows.Add();
        return new ExcelListRow(newRow);
    }

    #region IEnumerable<IExcelListRow> Support

    /// <summary>
    /// 返回枚举器，用于 foreach 遍历。
    /// </summary>
    /// <returns>枚举器</returns>
    public IEnumerator<IExcelListRow> GetEnumerator()
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelListRows));

        for (int i = 1; i <= _listRows.Count; i++)
        {
            yield return new ExcelListRow((MsExcel.ListRow)_listRows[i]);
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

        if (disposing)
        {
            if (_listRows != null)
            {
                Marshal.ReleaseComObject(_listRows);
                _listRows = null;
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 析构函数，确保即使未调用 Dispose 也能释放 COM 对象。
    /// </summary>
    ~ExcelListRows()
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