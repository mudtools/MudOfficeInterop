
using log4net;

namespace MudTools.OfficeInterop.Excel.Imps;

// =============================================
// 内部实现类：ExcelSlicerItems
// =============================================
internal class ExcelSlicerItems : IExcelSlicerItems
{
    internal MsExcel.SlicerItems _slicerItems;
    /// <summary>
    /// 用于记录此类型运行时日志的 logger 实例。
    /// </summary>
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelSlicerItems));
    private bool _disposedValue = false;

    /// <summary>
    /// 构造函数，初始化封装对象。
    /// </summary>
    /// <param name="slicerItems">原始的 COM SlicerItems 对象</param>
    internal ExcelSlicerItems(MsExcel.SlicerItems slicerItems)
    {
        _slicerItems = slicerItems ?? throw new ArgumentNullException(nameof(slicerItems));
    }

    /// <summary>
    /// 获取集合中切片器项的总数。
    /// </summary>
    public int Count => _slicerItems.Count;

    /// <summary>
    /// 通过索引（从 1 开始）或名称获取指定的切片器项。
    /// </summary>
    /// <param name="indexOrName">项索引（int）或名称（string）</param>
    /// <returns>对应的切片器项对象</returns>
    public IExcelSlicerItem this[object indexOrName]
    {
        get
        {
            if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelSlicerItems));
            try
            {
                return new ExcelSlicerItem(_slicerItems[indexOrName]);
            }
            catch (Exception ex)
            {
                log.Error($"获取切片器项（索引或名称：{indexOrName}）失败: {ex.Message}", ex);
                throw;
            }
        }
    }

    /// <summary>
    /// 获取此集合所属的父对象（通常是 Slicer 或 SlicerCache）。
    /// </summary>
    public object Parent => _slicerItems.Parent;

    /// <summary>
    /// 获取此集合所属的 Excel 应用程序对象。
    /// </summary>
    public IExcelApplication Application => new ExcelApplication(_slicerItems.Application);


    /// <summary>
    /// 选中所有项（显示所有数据）。
    /// </summary>
    public void SelectAll()
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelSlicerItems));
        try
        {
            foreach (MsExcel.SlicerItem item in _slicerItems)
            {
                item.Selected = true;
            }
        }
        catch (Exception ex)
        {
            log.Error($"选中所有切片器项失败: {ex.Message}", ex);
            throw;
        }
    }

    /// <summary>
    /// 取消选中所有项（隐藏所有数据，除非设置为允许空筛选）。
    /// </summary>
    public void UnselectAll()
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelSlicerItems));
        try
        {
            foreach (MsExcel.SlicerItem item in _slicerItems)
            {
                item.Selected = false;
            }
        }
        catch (Exception ex)
        {
            log.Error($"取消选中所有切片器项失败: {ex.Message}", ex);
            throw;
        }
    }

    #region IEnumerable<IExcelSlicerItem> Support

    /// <summary>
    /// 返回枚举器，用于 foreach 遍历。
    /// </summary>
    /// <returns>枚举器</returns>
    public IEnumerator<IExcelSlicerItem> GetEnumerator()
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelSlicerItems));

        for (int i = 1; i <= _slicerItems.Count; i++)
        {
            yield return new ExcelSlicerItem(_slicerItems[i]);
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
                if (_slicerItems != null)
                {
                    Marshal.ReleaseComObject(_slicerItems);
                    _slicerItems = null;
                }
            }
            catch (Exception ex)
            {
                log.Error($"释放 SlicerItems 时发生异常: {ex.Message}", ex);
                // 忽略释放异常，避免掩盖更严重的问题
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 析构函数，确保即使未调用 Dispose 也能释放 COM 对象。
    /// </summary>
    ~ExcelSlicerItems()
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