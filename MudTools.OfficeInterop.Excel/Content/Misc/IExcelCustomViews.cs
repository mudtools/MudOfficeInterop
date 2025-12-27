//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelCustomViews : IEnumerable<IExcelCustomView?>, IOfficeObject<IExcelCustomViews>, IDisposable
{
    /// <summary>
    /// 获取对象的父对象 
    /// 对应 IconSet.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取对象所在的Application对象
    /// 对应 IconSet.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取对象数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的列对象
    /// </summary>
    IExcelCustomView? this[int index] { get; }

    /// <summary>
    /// 获取指定名称的列对象
    /// </summary>
    IExcelCustomView? this[string viewName] { get; }

    /// <summary>
    /// 创建一个新的自定义视图。
    /// </summary>
    /// <param name="viewName">必需。新视图的名称。</param>
    /// <param name="printSettings">可选。True表示在自定义视图中包含打印设置。</param>
    /// <param name="rowColSettings">可选。True表示在自定义视图中包含隐藏行和列的设置（包括筛选信息）。</param>
    /// <returns>表示新视图的CustomView对象。</returns>
    IExcelCustomView? Add(string viewName, bool? printSettings = null, bool? rowColSettings = null);

}