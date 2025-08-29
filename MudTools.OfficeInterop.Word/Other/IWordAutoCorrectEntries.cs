namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示对 Microsoft Word AutoCorrectEntries 集合的封装接口。
/// 包含所有自动更正条目，支持按名称或索引访问、添加、删除等操作。
/// </summary>
public interface IWordAutoCorrectEntries : IEnumerable<IWordAutoCorrectEntry>, IDisposable
{
    /// <summary>
    /// 获取自动更正条目的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过 1-based 索引获取指定位置的自动更正条目。
    /// </summary>
    /// <param name="index">索引（从1开始）。</param>
    /// <returns>封装后的条目对象，若不存在则返回 null。</returns>
    IWordAutoCorrectEntry? this[int index] { get; }

    /// <summary>
    /// 通过名称（触发词）获取指定的自动更正条目。
    /// </summary>
    /// <param name="name">条目的名称（如 "teh"）。</param>
    /// <returns>封装后的条目对象，若不存在则返回 null。</returns>
    IWordAutoCorrectEntry? this[string name] { get; }

    /// <summary>
    /// 添加一个新的自动更正条目。
    /// </summary>
    /// <param name="name">触发词（如 "omw"）。</param>
    /// <param name="value">替换内容（如 "on my way"）。</param>
    /// <returns>新创建的封装条目对象。</returns>
    /// <exception cref="ArgumentException">当 name 或 value 为空时抛出。</exception>
    /// <exception cref="InvalidOperationException">当添加失败时抛出。</exception>
    IWordAutoCorrectEntry Add(string name, string value);

    /// <summary>
    /// 判断是否存在指定名称的自动更正条目。
    /// </summary>
    /// <param name="name">要查找的名称。</param>
    /// <returns>存在返回 true，否则 false。</returns>
    bool Contains(string name);

    /// <summary>
    /// 获取所有自动更正条目的名称列表。
    /// </summary>
    /// <returns>名称字符串列表。</returns>
    List<string> GetNames();

    /// <summary>
    /// 删除集合中的所有自动更正条目。
    /// </summary>
    void Clear();
}