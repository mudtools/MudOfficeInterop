//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 中的一个切片器缓存（Slicer Cache），用于管理切片器的数据源、筛选状态、排序等。
/// 一个切片器缓存可以被多个切片器控件共享，实现联动筛选。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelSlicerCache : IOfficeObject<IExcelSlicerCache>, IDisposable
{
    /// <summary>
    /// 获取此缓存所属的父对象（通常是 SlicerCaches 集合）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此缓存所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置此缓存的名称（全局唯一，用于绑定多个切片器）。
    /// 设置时若值为 null 将抛出 <see cref="ArgumentNullException"/>。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取此缓存关联的源名称（如透视表字段名或表格列名）。
    /// 若无源名称则返回空字符串。
    /// </summary>
    string SourceName { get; }

    /// <summary>
    /// 获取与此缓存关联的所有数据透视表集合。
    /// </summary>
    ISlicerPivotTables? PivotTables { get; }

    /// <summary>
    /// 获取与此缓存关联的所有切片器控件集合。
    /// </summary>
    IExcelSlicers? Slicers { get; }

    /// <summary>
    /// 获取此缓存中所有切片器项（包括隐藏和显示的）。
    /// </summary>
    IExcelSlicerItems? SlicerItems { get; }

    /// <summary>
    /// 获取当前可见（未被筛选掉）的切片器项集合。
    /// </summary>
    IExcelSlicerItems? VisibleSlicerItems { get; }

    /// <summary>
    /// 获取此缓存关联的列表对象（如 Excel 表格）。
    /// </summary>
    IExcelListObject? ListObject { get; }

    /// <summary>
    /// 获取此缓存的层级结构（用于多级切片器）。
    /// </summary>
    IExcelSlicerCacheLevels? SlicerCacheLevels { get; }

    /// <summary>
    /// 获取或设置切片器之间的交叉筛选类型。
    /// </summary>
    XlSlicerCrossFilterType CrossFilterType { get; set; }

    /// <summary>
    /// 获取或设置切片器项的排序方式（如升序、降序）。
    /// </summary>
    XlSlicerSort SortItems { get; set; }

    /// <summary>
    /// 获取此缓存的类型（如普通切片器、时间轴等）。
    /// </summary>
    XlSlicerCacheType SlicerCacheType { get; }

    /// <summary>
    /// 获取或设置是否使用自定义列表进行排序。
    /// </summary>
    bool SortUsingCustomLists { get; set; }

    /// <summary>
    /// 获取筛选器是否已被清除（即所有项恢复选中状态）。
    /// </summary>
    bool FilterCleared { get; }

    /// <summary>
    /// 获取此缓存是否关联到列表对象（非透视表）。
    /// </summary>
    bool List { get; }

    /// <summary>
    /// 获取是否需要手动更新筛选结果（适用于性能优化场景）。
    /// </summary>
    bool RequireManualUpdate { get; }

    /// <summary>
    /// 获取或设置是否显示所有项（即使未选中）。
    /// </summary>
    bool ShowAllItems { get; set; }

    /// <summary>
    /// 清除此缓存中所有项的手动筛选状态（恢复默认选中状态）。
    /// 若对象已被释放，将抛出 <see cref="ObjectDisposedException"/>。
    /// </summary>
    void ClearManualFilter();

    /// <summary>
    /// 清除所有筛选条件（包括手动和自动筛选）。
    /// 若对象已被释放，将抛出 <see cref="ObjectDisposedException"/>。
    /// </summary>
    void ClearAllFilters();

    /// <summary>
    /// 清除日期筛选条件（仅对日期类型字段有效）。
    /// 若对象已被释放，将抛出 <see cref="ObjectDisposedException"/>。
    /// </summary>
    void ClearDateFilter();

    /// <summary>
    /// 删除此切片器缓存（将同时删除所有关联的切片器控件）。
    /// 若对象已被释放，将抛出 <see cref="ObjectDisposedException"/>。
    /// 删除失败时会记录日志并重新抛出异常。
    /// </summary>
    void Delete();
}