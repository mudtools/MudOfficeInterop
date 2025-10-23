//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel PivotCaches 集合对象的二次封装实现类
/// 实现 IExcelPivotCaches 接口
/// </summary>
internal class ExcelPivotCaches : IExcelPivotCaches
{
    private MsExcel.PivotCaches _pivotCaches;
    private bool _disposedValue = false;

    internal ExcelPivotCaches(MsExcel.PivotCaches pivotCaches)
    {
        _pivotCaches = pivotCaches ?? throw new ArgumentNullException(nameof(pivotCaches));
    }

    #region 基础属性
    public int Count => _pivotCaches.Count;

    public IExcelPivotCache? this[int index]
    {
        get
        {
            if (_pivotCaches == null || index < 1 || index > Count)
                return null;

            try
            {
                var cacheObject = _pivotCaches[index];
                return cacheObject != null ? new ExcelPivotCache(cacheObject) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    public object Parent => _pivotCaches.Parent;

    public IExcelApplication Application => new ExcelApplication(_pivotCaches.Application);
    #endregion

    #region 创建和添加
    public IExcelPivotCache? Add(XlPivotTableSourceType sourceType, object sourceData)
    {
        if (_pivotCaches == null)
            return null;
        object comSourceData = GetSourceObj(sourceData);
        MsExcel.PivotCache newCache = _pivotCaches.Add(sourceType.EnumConvert(MsExcel.XlPivotTableSourceType.xlPivotTable), comSourceData);
        if (newCache != null)
            return new ExcelPivotCache(newCache);
        return null;
    }

    public IExcelPivotCache? Create(XlPivotTableSourceType sourceType, object? sourceData = null, object? version = null)
    {
        if (_pivotCaches == null)
            return null;
        // 处理可选参数
        object comVersion = version ?? System.Type.Missing;

        // sourceData 可能是字符串、Range、ListObject 等，直接传递给 Interop
        object comSourceData = GetSourceObj(sourceData);

        // 调用 Interop 方法
        MsExcel.PivotCache newCache = _pivotCaches.Create(
            sourceType.EnumConvert(MsExcel.XlPivotTableSourceType.xlPivotTable),
            comSourceData,
            comVersion
        );
        if (newCache != null)
            return new ExcelPivotCache(newCache);
        return null;
    }

    private object GetSourceObj(object? sourceData)
    {
        object comSourceData = Type.Missing;
        if (sourceData is ExcelRange rrange && rrange.InternalRange != null)
            comSourceData = rrange.InternalRange;
        else if (sourceData is ExcelListObject lo && lo._listObject != null)
            comSourceData = lo._listObject;
        else if (sourceData is ExcelPivotTable dt && dt._pivotTable != null)
            comSourceData = dt._pivotTable;
        else if (sourceData is string sourceString)
            comSourceData = sourceString;
        return comSourceData;
    }
    #endregion

    #region 查找和筛选   

    public IExcelPivotCache[] FindByMemoryUsage(long minSize)
    {
        var results = new List<IExcelPivotCache>();
        // 获取内存使用量比较复杂，需要特定API或估算
        // 占位符实现
        System.Diagnostics.Debug.WriteLine($"Finding caches by memory usage > {minSize} bytes is complex.");
        return results.ToArray();
    }
    #endregion

    #region 操作方法  

    public void Refresh()
    {
        RefreshAll();
    }
    #endregion

    #region 导出和导入
    public int ExportMetadataToFolder(string folderPath, string format = "json", string prefix = "pivotcache_")
    {
        if (!System.IO.Directory.Exists(folderPath)) return 0;

        int count = 0;
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                // var cache = this[i];
                // string info = cache.GetMetadataAsJson(); // 假设
                // string fileName = System.IO.Path.Combine(folderPath, $"{prefix}{cache.GetHashCode()}.{format}");
                // System.IO.File.WriteAllText(fileName, info);
                // 占位符实现
                string info = $" metadata for cache {i}";
                string fileName = System.IO.Path.Combine(folderPath, $"{prefix}{i}.{format}");
                System.IO.File.WriteAllText(fileName, info);
                count++;
            }
            catch
            {
                // Log error
            }
        }
        return count;
    }
    #endregion

    #region 高级功能

    public void RefreshAll()
    {
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                this[i].Refresh();
            }
            catch
            {

            }
        }
    }

    #endregion

    #region IEnumerable<IExcelPivotCache> Support
    public IEnumerator<IExcelPivotCache> GetEnumerator()
    {
        for (int i = 1; i <= _pivotCaches.Count; i++)
        {
            yield return new ExcelPivotCache(_pivotCaches.Item(i) as MsExcel.PivotCache);
        }
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                // 释放底层COM对象
                if (_pivotCaches != null)
                    Marshal.ReleaseComObject(_pivotCaches);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _pivotCaches = null;
        }
        _disposedValue = true;
    }

    ~ExcelPivotCaches()
    {

        Dispose(disposing: false);
    }

    public void Dispose()
    {

        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
    #endregion
}
