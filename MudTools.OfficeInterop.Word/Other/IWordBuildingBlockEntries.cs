namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示对 Microsoft Word 构建基块集合（BuildingBlockEntries）的封装接口。
/// 支持按索引或名称访问、添加、查询、遍历等操作。
/// </summary>
public interface IWordBuildingBlockEntries : IEnumerable<IWordBuildingBlock>, IDisposable
{
    /// <summary>
    /// 获取集合中构建基块的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取指定位置的构建基块（索引从1开始）。
    /// </summary>
    /// <param name="index">1-based 索引。</param>
    /// <returns>封装后的构建基块对象，若不存在则返回 null。</returns>
    IWordBuildingBlock? this[int index] { get; }

    /// <summary>
    /// 添加一个新的构建基块。
    /// </summary>
    /// <param name="name">构建基块的名称。</param>
    /// <param name="type"></param>
    /// <param name="category">所属类别。</param>
    /// <param name="value">内容文本。</param>
    /// <param name="insertOptions"></param>
    /// <returns>新创建的封装构建基块对象。</returns>
    /// <exception cref="ArgumentException">当 name 或 category 为空时抛出。</exception>
    /// <exception cref="InvalidOperationException">当添加失败时抛出。</exception>
    IWordBuildingBlock Add(string name, WdBuildingBlockTypes type, string category, string value, WdDocPartInsertOptions insertOptions);

    /// <summary>
    /// 获取集合中所有构建基块的名称列表。
    /// </summary>
    /// <returns>名称字符串列表。</returns>
    List<string> GetNames();

    /// <summary>
    /// 删除集合中的所有构建基块。
    /// </summary>
    void Clear();
}