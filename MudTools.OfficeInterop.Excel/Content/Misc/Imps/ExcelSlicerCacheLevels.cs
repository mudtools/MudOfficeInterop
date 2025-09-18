
namespace MudTools.OfficeInterop.Excel.Imps;


// =============================================
// 内部实现类：ExcelSlicerCacheLevels
// =============================================
internal class ExcelSlicerCacheLevels : IExcelSlicerCacheLevels
{
    internal MsExcel.SlicerCacheLevels _levels;
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelSlicerCacheLevels));
    private bool _disposedValue = false;

    /// <summary>
    /// 构造函数，初始化封装对象。
    /// </summary>
    /// <param name="levels">原始的 COM SlicerCacheLevels 对象</param>
    internal ExcelSlicerCacheLevels(MsExcel.SlicerCacheLevels levels)
    {
        _levels = levels ?? throw new ArgumentNullException(nameof(levels));
    }

    /// <summary>
    /// 获取集合中层级的总数。
    /// </summary>
    public int Count => _levels.Count;

    /// <summary>
    /// 通过索引（从 1 开始）获取指定的层级。
    /// </summary>
    /// <param name="index">层级索引（1-based）</param>
    /// <returns>对应的层级对象</returns>
    public IExcelSlicerCacheLevel this[int index]
    {
        get
        {
            if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelSlicerCacheLevels));
            try
            {
                return new ExcelSlicerCacheLevel(_levels[index]);
            }
            catch (Exception ex)
            {
                log.Error($"获取第 {index} 个层级失败: {ex.Message}", ex);
                throw;
            }
        }
    }

    /// <summary>
    /// 获取此集合所属的父对象（通常是 SlicerCache）。
    /// </summary>
    public object Parent => _levels.Parent;

    /// <summary>
    /// 获取此集合所属的 Excel 应用程序对象。
    /// </summary>
    public IExcelApplication Application => new ExcelApplication(_levels.Application);


    #region IEnumerable<IExcelSlicerCacheLevel> Support

    /// <summary>
    /// 返回枚举器，用于 foreach 遍历。
    /// </summary>
    /// <returns>枚举器</returns>
    public IEnumerator<IExcelSlicerCacheLevel> GetEnumerator()
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelSlicerCacheLevels));

        for (int i = 1; i <= _levels.Count; i++)
        {
            yield return new ExcelSlicerCacheLevel((MsExcel.SlicerCacheLevel)_levels[i]);
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
            try
            {
                if (_levels != null)
                {
                    Marshal.ReleaseComObject(_levels);
                    _levels = null;
                }
            }
            catch (Exception ex)
            {
                log.Error($"释放 SlicerCacheLevels 时发生异常: {ex.Message}", ex);
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 析构函数，确保即使未调用 Dispose 也能释放 COM 对象。
    /// </summary>
    ~ExcelSlicerCacheLevels()
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