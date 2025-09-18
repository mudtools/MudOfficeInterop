
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示切片器缓存中所有层级的集合，仅在 OLAP 数据源中有效。
/// 支持遍历和索引访问。
/// </summary>
public interface IExcelSlicerCacheLevels : IEnumerable<IExcelSlicerCacheLevel>, IDisposable
{
    /// <summary>
    /// 获取集合中层级的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引（从 1 开始）获取指定的层级。
    /// </summary>
    /// <param name="index">层级索引（1-based）</param>
    /// <returns>对应的层级对象</returns>
    IExcelSlicerCacheLevel this[int index] { get; }

    /// <summary>
    /// 获取此集合所属的父对象（通常是 SlicerCache）。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取此集合所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication Application { get; }
}