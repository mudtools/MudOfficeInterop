namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示对 Microsoft Word 中 Categories 集合的封装接口。
/// 该集合包含某一构建基块类型（如页眉、页脚）下的所有类别（Category）。
/// </summary>
public interface IWordCategories : IEnumerable<IWordCategory>, IDisposable
{
    /// <summary>
    /// 获取集合中类别的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过 1-based 索引获取指定位置的类别。
    /// </summary>
    /// <param name="index">索引（从1开始）。</param>
    /// <returns>封装后的类别对象，若不存在则返回 null。</returns>
    IWordCategory? this[int index] { get; }

    /// <summary>
    /// 通过名称获取指定的类别（区分大小写）。
    /// </summary>
    /// <param name="name">类别的名称。</param>
    /// <returns>封装后的类别对象，若不存在则返回 null。</returns>
    IWordCategory? this[string name] { get; }

    /// <summary>
    /// 判断是否存在指定名称的类别。
    /// </summary>
    /// <param name="name">要查找的类别名称。</param>
    /// <returns>存在返回 true，否则 false。</returns>
    bool Contains(string name);

    /// <summary>
    /// 获取所有类别的名称列表。
    /// </summary>
    /// <returns>类别名称字符串列表。</returns>
    List<string> GetNames();
}