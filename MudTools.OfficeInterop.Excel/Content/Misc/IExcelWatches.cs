
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定应用程序中所有 Watch 对象的集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelWatches : IEnumerable<IExcelWatch?>, IOfficeObject<IExcelWatches, MsExcel.Watches>, IDisposable
{
    /// <summary>
    /// 获取对象的父对象 
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取对象所在的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取对象数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引器获取集合中的指定 Watch 对象。
    /// </summary>
    /// <param name="index">对象的名称或索引号。</param>
    /// <returns>指定索引处的 Watch 对象。</returns>
    IExcelWatch? this[int index] { get; }

    /// <summary>
    /// 通过索引器获取集合中的指定 Watch 对象。
    /// </summary>
    /// <param name="name">对象的名称或索引号。</param>
    /// <returns>指定索引处的 Watch 对象。</returns>
    IExcelWatch? this[string name] { get; }

    /// <summary>
    /// 添加一个在重新计算工作表时跟踪的区域。返回一个 Watch 对象。
    /// </summary>
    /// <param name="source">区域的源。</param>
    /// <returns>新添加的 Watch 对象。</returns>
    IExcelWatch? Add(object source);

    /// <summary>
    /// 删除该对象集合中的所有监视对象。
    /// </summary>
    void Delete();
}