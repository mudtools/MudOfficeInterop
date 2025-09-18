
namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 表示切片器缓存中的一个层级（Level），仅在 OLAP 数据源中有效。
/// 用于处理多维数据（如年→季度→月）的层级筛选。
/// </summary>
public interface IExcelSlicerCacheLevel : IDisposable
{
    /// <summary>
    /// 获取此层级所属的父对象（通常是 SlicerCacheLevels 集合）。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取此层级所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取此层级在集合中的索引（从 1 开始）。
    /// </summary>
    int Index { get; }

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