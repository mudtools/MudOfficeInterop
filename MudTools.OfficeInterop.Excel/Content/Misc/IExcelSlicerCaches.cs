
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示工作簿中所有切片器缓存的集合，支持遍历、索引和名称访问。
/// </summary>
public interface IExcelSlicerCaches : IEnumerable<IExcelSlicerCache>, IDisposable
{
    /// <summary>
    /// 获取集合中切片器缓存的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引（从 1 开始）或名称获取指定的切片器缓存。
    /// </summary>
    /// <param name="indexOrName">缓存索引（int）或名称（string）</param>
    /// <returns>对应的切片器缓存对象</returns>
    IExcelSlicerCache this[object indexOrName] { get; }

    /// <summary>
    /// 获取此集合所属的父对象（通常是 Workbook）。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取此集合所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 创建一个新的切片器缓存并添加到集合中。
    /// </summary>
    /// <param name="source">数据源（PivotTable 或 ListObject 名称）</param>
    /// <param name="field">字段名称（透视表字段或表格列名）</param>
    /// <param name="name">缓存名称（可选，如不提供则自动生成）</param>
    /// <returns>新创建的切片器缓存对象</returns>
    IExcelSlicerCache Add(string source, string field, string name = null);
}