//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// 表示 Excel 切片器（Slicer）中的单个筛选项的封装接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.SlicerItem
/// 用于获取或设置切片器项的名称、值、选中状态等。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelSlicerItem : IOfficeObject<IExcelSlicerItem, MsExcel.SlicerItem>, IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 SlicerCache 或 SlicerItems 集合）。
    /// </summary>
    IExcelSlicerCache Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取切片器项的名称（显示在切片器 UI 上的文本）。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取切片器项的源名称（从数据源中获取的原始名称，可能需要转换）。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string SourceName { get; }

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