
namespace MudTools.OfficeInterop.Excel.Imps;
// =============================================
// 内部实现类：ExcelSlicerCaches
// =============================================
internal class ExcelSlicerCaches : IExcelSlicerCaches
{
    internal MsExcel.SlicerCaches _slicerCaches;
    private bool _disposedValue = false;
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelSlicerCaches));

    /// <summary>
    /// 构造函数，初始化封装对象。
    /// </summary>
    /// <param name="slicerCaches">原始的 COM SlicerCaches 对象</param>
    internal ExcelSlicerCaches(MsExcel.SlicerCaches slicerCaches)
    {
        _slicerCaches = slicerCaches ?? throw new ArgumentNullException(nameof(slicerCaches));
    }

    /// <summary>
    /// 获取集合中切片器缓存的总数。
    /// </summary>
    public int Count => _slicerCaches.Count;

    /// <summary>
    /// 通过索引（从 1 开始）或名称获取指定的切片器缓存。
    /// </summary>
    /// <param name="indexOrName">缓存索引（int）或名称（string）</param>
    /// <returns>对应的切片器缓存对象</returns>
    public IExcelSlicerCache this[object indexOrName]
    {
        get
        {
            if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelSlicerCaches));
            try
            {
                return new ExcelSlicerCache(_slicerCaches[indexOrName]);
            }
            catch (Exception ex)
            {
                log.Error($"获取切片器缓存（索引或名称：{indexOrName}）失败: {ex.Message}", ex);
                throw;
            }
        }
    }

    /// <summary>
    /// 获取此集合所属的父对象（通常是 Workbook）。
    /// </summary>
    public object Parent => _slicerCaches.Parent;

    /// <summary>
    /// 获取此集合所属的 Excel 应用程序对象。
    /// </summary>
    public IExcelApplication Application => new ExcelApplication(_slicerCaches.Application);

    /// <summary>
    /// 创建一个新的切片器缓存并添加到集合中。
    /// </summary>
    /// <param name="source">数据源（PivotTable 或 ListObject 名称）</param>
    /// <param name="field">字段名称（透视表字段或表格列名）</param>
    /// <param name="name">缓存名称（可选，如不提供则自动生成）</param>
    /// <returns>新创建的切片器缓存对象</returns>
    public IExcelSlicerCache Add(string source, string field, string name = null)
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelSlicerCaches));
        if (string.IsNullOrEmpty(source)) throw new ArgumentNullException(nameof(source));
        if (string.IsNullOrEmpty(field)) throw new ArgumentNullException(nameof(field));

        try
        {
            MsExcel.SlicerCache newCache;

            // 尝试作为透视表字段创建
            newCache = _slicerCaches.Add(
                Source: source,
                SourceField: field,
                Name: name
            );

            return new ExcelSlicerCache(newCache);
        }
        catch (Exception ex)
        {
            log.Error($"添加切片器缓存（源：{source}, 字段：{field}）失败: {ex.Message}", ex);
            throw;
        }
    }

    #region IEnumerable<IExcelSlicerCache> Support

    /// <summary>
    /// 返回枚举器，用于 foreach 遍历。
    /// </summary>
    /// <returns>枚举器</returns>
    public IEnumerator<IExcelSlicerCache> GetEnumerator()
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelSlicerCaches));

        for (int i = 1; i <= _slicerCaches.Count; i++)
        {
            yield return new ExcelSlicerCache(_slicerCaches[i]);
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
                if (_slicerCaches != null)
                {
                    Marshal.ReleaseComObject(_slicerCaches);
                    _slicerCaches = null;
                }
            }
            catch (Exception ex)
            {
                log.Error($"释放 SlicerCaches 时发生异常: {ex.Message}", ex);
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 析构函数，确保即使未调用 Dispose 也能释放 COM 对象。
    /// </summary>
    ~ExcelSlicerCaches()
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