//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 表示切片器缓存中的一个层级（Level），仅在 OLAP 数据源中有效。
/// 用于处理多维数据（如年→季度→月）的层级筛选。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelSlicerCacheLevel : IDisposable
{
    /// <summary>
    /// 获取此层级所属的父对象（通常是 SlicerCacheLevels 集合）。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取此层级所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取此层级在其父集合中的序号位置（从 1 开始计数）。
    /// </summary>
    int Ordinal { get; }

    /// <summary>
    /// 获取当前层级中项目的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取或设置切片器项目的排序方式。
    /// </summary>
    XlSlicerSort SortItems { get; set; }

    /// <summary>
    /// 获取或设置交叉筛选类型，用于控制切片器在交叉筛选时的行为。
    /// </summary>
    XlSlicerCrossFilterType CrossFilterType { get; set; }

    /// <summary>
    /// 获取此层级的名称（如“年”、“季度”、“产品类别”等）。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取此层级中所有切片器项的集合（仅在 OLAP 模式下可用）。
    /// </summary>
    IExcelSlicerItems SlicerItems { get; }
}