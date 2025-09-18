
namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// 表示 Excel 切片器（Slicer）中的单个筛选项的封装接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.SlicerItem
/// 用于获取或设置切片器项的名称、值、选中状态等。
/// </summary>
public interface IExcelSlicerItem : IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 SlicerCache 或 SlicerItems 集合）。
    /// </summary>
    IExcelSlicerCache Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取切片器项的名称（显示在切片器 UI 上的文本）。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取切片器项的标准化源名称（通常是从数据源中获取的原始名称）。
    /// </summary>
    string SourceNameStandard { get; }

    /// <summary>
    /// 获取或设置切片器项的显示标题（用于在UI上显示的文本）。
    /// </summary>
    string Caption { get; }

    /// <summary>
    /// 获取切片器项的实际值（可能与显示名称不同）。
    /// </summary>
    string Value { get; }

    /// <summary>
    /// 获取或设置该切片器项是否被选中。
    /// true = 选中（参与筛选），false = 取消选中（被过滤）。
    /// </summary>
    bool Selected { get; set; }

    /// <summary>
    /// 获取该切片器项是否具有数据（即是否关联到数据源中的记录）。
    /// false 表示该项无数据，通常显示为灰色不可选。
    /// </summary>
    bool HasData { get; }
}