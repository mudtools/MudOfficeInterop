//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

// =============================================
// 内部实现类：ExcelListRow
// =============================================
internal class ExcelListRow : IExcelListRow
{
    internal MsExcel.ListRow _listRow;
    private bool _disposedValue = false;

    /// <summary>
    /// 构造函数，初始化封装对象。
    /// </summary>
    /// <param name="listRow">原始的 COM ListRow 对象</param>
    internal ExcelListRow(MsExcel.ListRow listRow)
    {
        _listRow = listRow ?? throw new ArgumentNullException(nameof(listRow));
    }

    /// <summary>
    /// 获取此行所属的父对象（通常是 ListObject）。
    /// </summary>
    public object Parent => _listRow.Parent;

    /// <summary>
    /// 获取此行所属的 Excel 应用程序对象。
    /// </summary>
    public IExcelApplication Application => new ExcelApplication(_listRow.Application);

    /// <summary>
    /// 获取此行在 ListRows 集合中的索引（从 1 开始）。
    /// </summary>
    public int Index => _listRow.Index;

    public bool InvalidData => _listRow.InvalidData;

    /// <summary>
    /// 获取此行对应的单元格范围（Range），包含该行所有列的数据。
    /// </summary>
    public IExcelRange Range
    {
        get
        {
            if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelListRow));
            var range = _listRow.Range;
            return range != null ? new ExcelRange(range) : null;
        }
    }

    /// <summary>
    /// 删除此行（将从表格中移除该数据行）。
    /// </summary>
    public void Delete()
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelListRow));
        _listRow.Delete();
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
            if (_listRow != null)
            {
                Marshal.ReleaseComObject(_listRow);
                _listRow = null;
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 析构函数，确保即使未调用 Dispose 也能释放 COM 对象。
    /// </summary>
    ~ExcelListRow()
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