
namespace MudTools.OfficeInterop.Excel.Imps;


// =============================================
// 内部实现类：ExcelSlicers
// =============================================
internal class ExcelSlicers : IExcelSlicers
{
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelSlicerItems));
    internal MsExcel.Slicers _slicers;
    private bool _disposedValue = false;

    /// <summary>
    /// 构造函数，初始化封装对象。
    /// </summary>
    /// <param name="slicers">原始的 COM Slicers 对象</param>
    internal ExcelSlicers(MsExcel.Slicers slicers)
    {
        _slicers = slicers ?? throw new ArgumentNullException(nameof(slicers));
    }

    /// <summary>
    /// 获取集合中切片器的总数。
    /// </summary>
    public int Count => _slicers.Count;

    /// <summary>
    /// 通过索引（从 1 开始）或名称获取指定的切片器。
    /// </summary>
    /// <param name="indexOrName">切片器索引（int）或名称（string）</param>
    /// <returns>对应的切片器对象</returns>
    public IExcelSlicer this[object indexOrName]
    {
        get
        {
            if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelSlicers));
            try
            {
                return new ExcelSlicer((MsExcel.Slicer)_slicers[indexOrName]);
            }
            catch (Exception ex)
            {
                log.Error($"获取切片器（索引或名称：{indexOrName}）失败: {ex.Message}", ex);
                throw;
            }
        }
    }

    /// <summary>
    /// 获取此集合所属的父对象（通常是 Worksheet）。
    /// </summary>
    public object Parent => _slicers.Parent;

    /// <summary>
    /// 获取此集合所属的 Excel 应用程序对象。
    /// </summary>
    public IExcelApplication Application => new ExcelApplication(_slicers.Application);

    /// <summary>
    /// 向集合中添加一个新切片器。
    /// </summary>
    /// <param name="slicerCache">切片器缓存名称（必须已存在）</param>
    /// <param name="name">切片器名称（可选）</param>
    /// <param name="caption">切片器标题（可选）</param>
    /// <param name="top">距离工作表顶部的位置（点，可选）</param>
    /// <param name="left">距离工作表左侧的位置（点，可选）</param>
    /// <returns>新创建的切片器对象</returns>
    public IExcelSlicer Add(
        string slicerCache,
        string name = null,
        string caption = null,
        double? top = null,
        double? left = null)
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelSlicers));
        if (string.IsNullOrEmpty(slicerCache)) throw new ArgumentNullException(nameof(slicerCache));

        try
        {
            MsExcel.Slicer newSlicer;

            if (!string.IsNullOrEmpty(name))
            {
                newSlicer = _slicers.Add(slicerCache, Name: name);
            }
            else
            {
                newSlicer = _slicers.Add(slicerCache);
            }

            // 设置可选属性
            if (!string.IsNullOrEmpty(caption))
                newSlicer.Caption = caption;

            if (top.HasValue && left.HasValue)
            {
                newSlicer.Top = top.Value;
                newSlicer.Left = left.Value;
            }
            else if (top.HasValue || left.HasValue)
            {
                // 如果只设置一个，需先获取当前值
                double currentTop = newSlicer.Top;
                double currentLeft = newSlicer.Left;
                newSlicer.Top = top ?? currentTop;
                newSlicer.Left = left ?? currentLeft;
            }

            return new ExcelSlicer(newSlicer);
        }
        catch (Exception ex)
        {
            log.Error($"添加切片器（缓存：{slicerCache}）失败: {ex.Message}", ex);
            throw;
        }
    }

    #region IEnumerable<IExcelSlicer> Support

    /// <summary>
    /// 返回枚举器，用于 foreach 遍历。
    /// </summary>
    /// <returns>枚举器</returns>
    public IEnumerator<IExcelSlicer> GetEnumerator()
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelSlicers));

        for (int i = 1; i <= _slicers.Count; i++)
        {
            yield return new ExcelSlicer((MsExcel.Slicer)_slicers[i]);
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
                if (_slicers != null)
                {
                    Marshal.ReleaseComObject(_slicers);
                    _slicers = null;
                }
            }
            catch (Exception ex)
            {
                log.Error($"释放 Slicers 时发生异常: {ex.Message}", ex);
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 析构函数，确保即使未调用 Dispose 也能释放 COM 对象。
    /// </summary>
    ~ExcelSlicers()
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