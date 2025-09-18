namespace MudTools.OfficeInterop.Excel.Imps;


/// <summary>
/// TableObject（实际为 ListObject）COM 对象的封装实现类。
/// 负责管理 COM 对象生命周期，提供安全的属性访问和资源释放。
/// </summary>
internal class ExcelTableObject : IExcelTableObject
{
    /// <summary>
    /// 内部持有的原始 COM 对象（实际是 ListObject）。
    /// </summary>
    internal MsExcel.TableObject _listObject;

    /// <summary>
    /// 标记对象是否已被释放。
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装类。
    /// </summary>
    /// <param name="listObject">原始的 ListObject COM 对象，不可为 null。</param>
    /// <exception cref="ArgumentNullException">当传入的 listObject 为 null 时抛出。</exception>
    internal ExcelTableObject(MsExcel.TableObject listObject)
    {
        _listObject = listObject ?? throw new ArgumentNullException(nameof(listObject));
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
            if (_listObject != null)
            {
                Marshal.ReleaseComObject(_listObject);
                _listObject = null;
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
    /// 获取此对象的父对象（通常是 Worksheet）。
    /// </summary>
    public object? Parent => _listObject?.Parent;

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// 返回封装后的 <see cref="IExcelApplication"/> 接口实例。
    /// </summary>
    public IExcelApplication? Application =>
        _listObject?.Application != null
            ? new ExcelApplication(_listObject.Application as MsExcel.Application)
            : null;

    public bool RowNumbers
    {
        get => _listObject != null && _listObject.RowNumbers;
        set
        {
            if (_listObject != null)
                _listObject.RowNumbers = value;
        }
    }

    public bool EnableRefresh
    {
        get => _listObject != null && _listObject.EnableRefresh;
        set
        {
            if (_listObject != null)
                _listObject.EnableRefresh = value;
        }
    }

    public XlCellInsertionMode RefreshStyle
    {
        get => _listObject != null ? _listObject.RefreshStyle.EnumConvert(XlCellInsertionMode.xlOverwriteCells) : XlCellInsertionMode.xlOverwriteCells;
        set
        {
            if (_listObject != null)
                _listObject.RefreshStyle = value.EnumConvert(MsExcel.XlCellInsertionMode.xlOverwriteCells);
        }
    }

    public bool FetchedRowOverflow
    {
        get => _listObject != null && _listObject.FetchedRowOverflow;
    }

    public bool EnableEditing
    {
        get => _listObject != null && _listObject.EnableEditing;
        set
        {
            if (_listObject != null)
                _listObject.EnableEditing = value;
        }
    }

    public bool PreserveColumnInfo
    {
        get => _listObject != null && _listObject.PreserveColumnInfo;
        set
        {
            if (_listObject != null)
                _listObject.PreserveColumnInfo = value;
        }
    }

    public bool PreserveFormatting
    {
        get => _listObject != null && _listObject.PreserveFormatting;
        set
        {
            if (_listObject != null)
                _listObject.PreserveFormatting = value;
        }

    }

    public bool AdjustColumnWidth
    {
        get => _listObject != null && _listObject.AdjustColumnWidth;
        set
        {
            if (_listObject != null)
                _listObject.AdjustColumnWidth = value;
        }
    }

    public IExcelRange? Destination =>
        _listObject?.Destination != null
            ? new ExcelRange(_listObject.Destination)
            : null;

    public IExcelRange? ResultRange
    {
        get =>
            _listObject?.ResultRange != null
                ? new ExcelRange(_listObject.ResultRange)
                : null;
    }

    public IExcelListObject? ListObject
    {
        get =>
            _listObject?.ListObject != null
                ? new ExcelListObject(_listObject.ListObject)
                : null;
    }

    public IExcelWorkbookConnection? WorkbookConnection
    {
        get =>
            _listObject?.WorkbookConnection != null
                ? new ExcelWorkbookConnection(_listObject.WorkbookConnection)
                : null;
    }

    /// <summary>
    /// 删除整个表格（保留数据，移除表格结构）。
    /// </summary>
    public void Delete()
    {
        _listObject?.Delete();
    }

    public bool Refresh()
    {
        return _listObject?.Refresh() ?? false;
    }
}