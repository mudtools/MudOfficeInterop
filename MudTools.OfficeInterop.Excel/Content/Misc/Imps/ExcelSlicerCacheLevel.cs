
namespace MudTools.OfficeInterop.Excel.Imps;

// =============================================
// 内部实现类：ExcelSlicerCacheLevel
// =============================================
internal class ExcelSlicerCacheLevel : IExcelSlicerCacheLevel
{
    internal MsExcel.SlicerCacheLevel _level;
    private bool _disposedValue = false;

    /// <summary>
    /// 构造函数，初始化封装对象。
    /// </summary>
    /// <param name="level">原始的 COM SlicerCacheLevel 对象</param>
    internal ExcelSlicerCacheLevel(MsExcel.SlicerCacheLevel level)
    {
        _level = level ?? throw new ArgumentNullException(nameof(level));
    }

    /// <summary>
    /// 获取此层级所属的父对象（通常是 SlicerCacheLevels 集合）。
    /// </summary>
    public object Parent => _level.Parent;

    /// <summary>
    /// 获取此层级所属的 Excel 应用程序对象。
    /// </summary>
    public IExcelApplication Application => new ExcelApplication(_level.Application);

    /// <summary>
    /// 获取此层级在集合中的索引（从 1 开始）。
    /// </summary>
    public int Index => _level.Ordinal;

    public int Count => _level.Count;

    /// <summary>
    /// 获取此层级的名称（如“年”、“季度”、“产品类别”等）。
    /// </summary>
    public string Name => _level.Name;

    public IExcelSlicerItems SlicerItems => _level != null ? new ExcelSlicerItems(_level.SlicerItems) : null;

    public XlSlicerSort SortItems
    {
        get => _level.SortItems.EnumConvert(XlSlicerSort.xlSlicerSortAscending);
        set
        {
            if (_level != null)
                _level.SortItems = value.EnumConvert(MsExcel.XlSlicerSort.xlSlicerSortAscending);
        }
    }

    public XlSlicerCrossFilterType CrossFilterType
    {
        get => _level.CrossFilterType.EnumConvert(XlSlicerCrossFilterType.xlSlicerNoCrossFilter);
        set
        {
            if (_level != null)
                _level.CrossFilterType = value.EnumConvert(MsExcel.XlSlicerCrossFilterType.xlSlicerNoCrossFilter);
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
            if (_level != null)
            {
                Marshal.ReleaseComObject(_level);
                _level = null;
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 析构函数，确保即使未调用 Dispose 也能释放 COM 对象。
    /// </summary>
    ~ExcelSlicerCacheLevel()
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