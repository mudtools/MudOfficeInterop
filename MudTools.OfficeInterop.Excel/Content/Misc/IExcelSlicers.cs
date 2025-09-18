
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示工作表中所有切片器的集合，支持遍历和索引访问。
/// </summary>
public interface IExcelSlicers : IEnumerable<IExcelSlicer>, IDisposable
{
    /// <summary>
    /// 获取集合中切片器的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引（从 1 开始）或名称获取指定的切片器。
    /// </summary>
    /// <param name="indexOrName">切片器索引（int）或名称（string）</param>
    /// <returns>对应的切片器对象</returns>
    IExcelSlicer this[object indexOrName] { get; }

    /// <summary>
    /// 获取此集合所属的父对象（通常是 Worksheet）。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取此集合所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 向集合中添加一个新切片器。
    /// </summary>
    /// <param name="slicerCache">切片器缓存名称（必须已存在）</param>
    /// <param name="name">切片器名称（可选）</param>
    /// <param name="caption">切片器标题（可选）</param>
    /// <param name="top">距离工作表顶部的位置（点，可选）</param>
    /// <param name="left">距离工作表左侧的位置（点，可选）</param>
    /// <returns>新创建的切片器对象</returns>
    IExcelSlicer Add(
        string slicerCache,
        string name = null,
        string caption = null,
        double? top = null,
        double? left = null);
}