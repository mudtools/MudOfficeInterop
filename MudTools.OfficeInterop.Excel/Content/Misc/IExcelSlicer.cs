namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 中的一个切片器对象，用于对数据透视表或表格进行可视化筛选。
/// </summary>
public interface IExcelSlicer : IDisposable
{
    /// <summary>
    /// 获取此切片器所属的父对象（通常是 Slicers 集合）。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取此切片器所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置切片器的名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置切片器的标题。
    /// </summary>
    string Caption { get; set; }

    /// <summary>
    /// 获取切片器关联的字段名称（即筛选依据的列名）。
    /// </summary>
    IExcelSlicerCache? SlicerCache { get; }

    /// <summary>
    /// 获取切片器的形状对象，可用于设置切片器的位置、样式等外观属性。
    /// </summary>
    IExcelShape? Shape { get; }

    /// <summary>
    /// 获取切片器当前活动的项，表示用户当前选中或高亮显示的切片器项。
    /// </summary>
    IExcelSlicerItem? ActiveItem { get; }

    /// <summary>
    /// 获取切片器缓存的层级信息，包含切片器项的排序方式、交叉筛选类型等设置。
    /// </summary>
    IExcelSlicerCacheLevel? SlicerCacheLevel { get; }

    /// <summary>
    /// 获取或设置切片器的宽度（点，points）。
    /// </summary>
    double Width { get; set; }

    /// <summary>
    /// 获取或设置切片器的高度（点，points）。
    /// </summary>
    double Height { get; set; }

    /// <summary>
    /// 获取或设置切片器标题是否可见。
    /// </summary>
    bool DisplayHeader { get; set; }

    /// <summary>
    /// 获取或设置切片器项的列数。
    /// </summary>
    int Columns { get; set; }

    /// <summary>
    /// 获取切片器缓存的类型，表示切片器是标准切片器还是时间线切片器。
    /// </summary>
    XlSlicerCacheType SlicerCacheType { get; }

    /// <summary>
    /// 删除此切片器。
    /// </summary>
    void Delete();

    /// <summary>
    /// 清除切片器的所有筛选状态（显示所有项）。
    /// </summary>
    void ClearManualFilter();
}