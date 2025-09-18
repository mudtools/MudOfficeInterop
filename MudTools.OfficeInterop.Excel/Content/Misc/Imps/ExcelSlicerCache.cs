
namespace MudTools.OfficeInterop.Excel.Imps;
// =============================================
// 内部实现类：ExcelSlicerCache
// =============================================
internal class ExcelSlicerCache : IExcelSlicerCache
{
    internal MsExcel.SlicerCache _slicerCache;
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelSlicerCache));
    private bool _disposedValue = false;

    /// <summary>
    /// 构造函数，初始化封装对象。
    /// </summary>
    /// <param name="slicerCache">原始的 COM SlicerCache 对象</param>
    internal ExcelSlicerCache(MsExcel.SlicerCache slicerCache)
    {
        _slicerCache = slicerCache ?? throw new ArgumentNullException(nameof(slicerCache));
    }

    /// <summary>
    /// 获取此缓存所属的父对象（通常是 SlicerCaches 集合）。
    /// </summary>
    public object Parent => _slicerCache.Parent;

    /// <summary>
    /// 获取此缓存所属的 Excel 应用程序对象。
    /// </summary>
    public IExcelApplication Application => new ExcelApplication(_slicerCache.Application);

    /// <summary>
    /// 获取或设置此缓存的名称（全局唯一，用于绑定多个切片器）。
    /// </summary>
    public string Name
    {
        get => _slicerCache.Name;
        set => _slicerCache.Name = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>
    /// 获取此缓存关联的源名称（如透视表字段名或表格列名）。
    /// </summary>
    public string SourceName => _slicerCache.SourceName?.ToString() ?? string.Empty;

    /// <summary>
    /// 获取此缓存是否基于数据透视表。
    /// </summary>
    public bool IsPivot => _slicerCache.PivotTables?.Count > 0;

    public ISlicerPivotTables PivotTables => _slicerCache != null ? new SlicerPivotTables(_slicerCache.PivotTables) : null;

    public IExcelSlicers Slicers => _slicerCache != null ? new ExcelSlicers(_slicerCache.Slicers) : null;

    public IExcelSlicerItems SlicerItems => _slicerCache != null ? new ExcelSlicerItems(_slicerCache.SlicerItems) : null;

    public IExcelSlicerItems VisibleSlicerItems => _slicerCache != null ? new ExcelSlicerItems(_slicerCache.VisibleSlicerItems) : null;

    public IExcelListObject ListObject => _slicerCache != null ? new ExcelListObject(_slicerCache.ListObject) : null;

    public IExcelSlicerCacheLevels SlicerCacheLevels => _slicerCache != null ? new ExcelSlicerCacheLevels(_slicerCache.SlicerCacheLevels) : null;

    public XlSlicerCrossFilterType CrossFilterType
    {
        get => _slicerCache.CrossFilterType.EnumConvert(XlSlicerCrossFilterType.xlSlicerNoCrossFilter);
        set
        {
            if (_slicerCache != null)
                _slicerCache.CrossFilterType = value.EnumConvert(MsExcel.XlSlicerCrossFilterType.xlSlicerNoCrossFilter);
        }
    }

    public XlSlicerSort SortItems
    {
        get => _slicerCache.SortItems.EnumConvert(XlSlicerSort.xlSlicerSortAscending);
        set
        {
            if (_slicerCache != null)
                _slicerCache.SortItems = value.EnumConvert(MsExcel.XlSlicerSort.xlSlicerSortAscending);
        }
    }

    public XlSlicerCacheType SlicerCacheType
    {
        get => _slicerCache.SlicerCacheType.EnumConvert(XlSlicerCacheType.xlSlicer);
    }

    public bool SortUsingCustomLists
    {
        get => _slicerCache.SortUsingCustomLists;
        set
        {
            if (_slicerCache != null)
                _slicerCache.SortUsingCustomLists = value;
        }
    }

    public bool FilterCleared
    {
        get => _slicerCache.FilterCleared;
    }

    public bool List
    {
        get => _slicerCache.List;
    }

    public bool RequireManualUpdate
    {
        get => _slicerCache.RequireManualUpdate;
    }

    public bool ShowAllItems
    {
        get => _slicerCache.ShowAllItems;
        set
        {
            if (_slicerCache != null)
                _slicerCache.ShowAllItems = value;
        }

    }

    /// <summary>
    /// 清除此缓存中所有项的手动筛选状态（恢复默认选中状态）。
    /// </summary>
    public void ClearManualFilter()
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelSlicerCache));
        try
        {
            _slicerCache.ClearManualFilter();
        }
        catch (Exception ex)
        {
            log.Error($"清除缓存 '{this.Name}' 的筛选状态失败: {ex.Message}", ex);
        }
    }

    public void ClearAllFilters()
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelSlicerCache));
        try
        {
            _slicerCache.ClearAllFilters();
        }
        catch (Exception ex)
        {
            log.Error($"清除缓存 '{this.Name}' 的筛选状态失败: {ex.Message}", ex);
        }
    }

    public void ClearDateFilter()
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelSlicerCache));
        try
        {
            _slicerCache.ClearDateFilter();
        }
        catch (Exception ex)
        {
            log.Error($"清除缓存 '{this.Name}' 的筛选状态失败: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 删除此切片器缓存（将同时删除所有关联的切片器控件）。
    /// </summary>
    public void Delete()
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelSlicerCache));
        try
        {
            _slicerCache.Delete();
        }
        catch (Exception ex)
        {
            log.Error($"删除缓存 '{this.Name}' 失败: {ex.Message}", ex);
            throw;
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
            if (_slicerCache != null)
            {
                Marshal.ReleaseComObject(_slicerCache);
                _slicerCache = null;
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 析构函数，确保即使未调用 Dispose 也能释放 COM 对象。
    /// </summary>
    ~ExcelSlicerCache()
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