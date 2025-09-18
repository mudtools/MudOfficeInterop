
namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// SlicerItem COM 对象的封装实现类。
/// 负责管理 COM 对象生命周期，提供安全的属性访问和资源释放。
/// </summary>
internal class ExcelSlicerItem : IExcelSlicerItem
{
    /// <summary>
    /// 内部持有的原始 COM 对象。
    /// </summary>
    internal MsExcel.SlicerItem _slicerItem;

    /// <summary>
    /// 标记对象是否已被释放。
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装类。
    /// </summary>
    /// <param name="slicerItem">原始的 SlicerItem COM 对象，不可为 null。</param>
    /// <exception cref="ArgumentNullException">当传入的 slicerItem 为 null 时抛出。</exception>
    internal ExcelSlicerItem(MsExcel.SlicerItem slicerItem)
    {
        _slicerItem = slicerItem ?? throw new ArgumentNullException(nameof(slicerItem));
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
            if (_slicerItem != null)
            {
                Marshal.ReleaseComObject(_slicerItem);
                _slicerItem = null;
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
    /// 获取此对象的父对象（通常是 SlicerItems 集合）。
    /// </summary>
    public IExcelSlicerCache Parent => _slicerItem != null ? new ExcelSlicerCache(_slicerItem?.Parent) : null;

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// 返回封装后的 <see cref="IExcelApplication"/> 接口实例。
    /// </summary>
    public IExcelApplication Application =>
        _slicerItem?.Application != null
            ? new ExcelApplication(_slicerItem.Application as MsExcel.Application)
            : null;

    /// <summary>
    /// 获取切片器项的名称（显示在切片器 UI 上的文本）。
    /// </summary>
    public string Name => _slicerItem?.Name ?? string.Empty;

    /// <summary>
    /// 获取切片器项的实际值（可能与显示名称不同）。
    /// </summary>
    public string Value => _slicerItem?.Value ?? string.Empty;

    public string SourceNameStandard => _slicerItem?.SourceNameStandard ?? string.Empty;

    public string Caption => _slicerItem?.Caption ?? string.Empty;

    /// <summary>
    /// 获取或设置该切片器项是否被选中。
    /// 设置为 true 时，该项参与筛选；false 时被排除。
    /// 注意：如果 HasData=false，设置 Selected 可能无效。
    /// </summary>
    public bool Selected
    {
        get => _slicerItem != null && _slicerItem.Selected;
        set
        {
            if (_slicerItem != null)
                _slicerItem.Selected = value;
        }
    }

    /// <summary>
    /// 获取该切片器项是否具有数据（即是否关联到数据源中的记录）。
    /// false 表示该项无数据，通常显示为灰色不可选。
    /// </summary>
    public bool HasData => _slicerItem != null && _slicerItem.HasData;


}