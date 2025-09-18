
namespace MudTools.OfficeInterop.Excel.Imps;

// =============================================
// 内部实现类：ExcelSlicer
// =============================================
internal class ExcelSlicer : IExcelSlicer
{
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelSlicerItems));
    internal MsExcel.Slicer _slicer;
    private bool _disposedValue = false;

    /// <summary>
    /// 构造函数，初始化封装对象。
    /// </summary>
    /// <param name="slicer">原始的 COM Slicer 对象</param>
    internal ExcelSlicer(MsExcel.Slicer slicer)
    {
        _slicer = slicer ?? throw new ArgumentNullException(nameof(slicer));
    }

    /// <summary>
    /// 获取此切片器所属的父对象（通常是 Slicers 集合）。
    /// </summary>
    public object Parent => _slicer.Parent;

    /// <summary>
    /// 获取此切片器所属的 Excel 应用程序对象。
    /// </summary>
    public IExcelApplication Application => new ExcelApplication(_slicer.Application);

    /// <summary>
    /// 获取或设置切片器的名称。
    /// </summary>
    public string Name
    {
        get => _slicer.Name;
        set => _slicer.Name = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>
    /// 获取或设置切片器的标题。
    /// </summary>
    public string Caption
    {
        get => _slicer.Caption;
        set => _slicer.Caption = value ?? throw new ArgumentNullException(nameof(value));
    }

    public IExcelShape? Shape => _slicer != null ? new ExcelShape(_slicer.Shape) : null;

    public IExcelSlicerItem? ActiveItem => _slicer != null ? new ExcelSlicerItem(_slicer.ActiveItem) : null;

    public IExcelSlicerCache? SlicerCache => _slicer != null ? new ExcelSlicerCache(_slicer.SlicerCache) : null;

    public IExcelSlicerCacheLevel? SlicerCacheLevel => _slicer != null ? new ExcelSlicerCacheLevel(_slicer.SlicerCacheLevel) : null;


    /// <summary>
    /// 获取或设置切片器的宽度（点，points）。
    /// </summary>
    public double Width
    {
        get => _slicer.Width;
        set => _slicer.Width = value;
    }

    /// <summary>
    /// 获取或设置切片器的高度（点，points）。
    /// </summary>
    public double Height
    {
        get => _slicer.Height;
        set => _slicer.Height = value;
    }

    /// <summary>
    /// 获取或设置切片器标题是否可见。
    /// </summary>
    public bool DisplayHeader
    {
        get => _slicer.DisplayHeader;
        set => _slicer.DisplayHeader = value;
    }

    /// <summary>
    /// 获取或设置切片器项的列数。
    /// </summary>
    public int Columns
    {
        get => _slicer.NumberOfColumns;
        set => _slicer.NumberOfColumns = value;
    }

    public XlSlicerCacheType SlicerCacheType
    {
        get => _slicer.SlicerCacheType.EnumConvert(XlSlicerCacheType.xlSlicer);
    }

    /// <summary>
    /// 删除此切片器。
    /// </summary>
    public void Delete()
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelSlicer));
        try
        {
            _slicer.Delete();
        }
        catch (Exception ex)
        {
            log.Error($"删除切片器 '{this.Name}' 失败: {ex.Message}", ex);
            throw;
        }
    }

    /// <summary>
    /// 清除切片器的所有筛选状态（显示所有项）。
    /// </summary>
    public void ClearManualFilter()
    {
        if (_disposedValue) throw new ObjectDisposedException(nameof(ExcelSlicer));
        try
        {
            _slicer.SlicerCache.ClearManualFilter();
        }
        catch (Exception ex)
        {
            log.Error($"清除切片器 '{this.Name}' 筛选状态失败: {ex.Message}", ex);
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
            if (_slicer != null)
            {
                Marshal.ReleaseComObject(_slicer);
                _slicer = null;
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 析构函数，确保即使未调用 Dispose 也能释放 COM 对象。
    /// </summary>
    ~ExcelSlicer()
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