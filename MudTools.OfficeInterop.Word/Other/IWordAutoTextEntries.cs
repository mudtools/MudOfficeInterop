namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 中“自动图文集条目”集合的封装接口。
/// 提供对模板中所有自动图文集条目的访问、添加和枚举功能。
/// </summary>
public interface IWordAutoTextEntries : IEnumerable<IWordAutoTextEntry>, IDisposable
{
    /// <summary>
    /// 获取集合中自动图文集条目的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取指定的自动图文集条目（从 1 开始）
    /// </summary>
    /// <param name="index">索引（1-based）</param>
    /// <returns>封装后的条目对象</returns>
    IWordAutoTextEntry? this[int index] { get; }

    /// <summary>
    /// 根据名称获取指定的自动图文集条目
    /// </summary>
    /// <param name="name">条目名称</param>
    /// <returns>封装后的条目对象；若不存在返回 null</returns>
    IWordAutoTextEntry? this[string name] { get; }

    /// <summary>
    /// 向模板中添加一个新的自动图文集条目
    /// </summary>
    /// <param name="name">条目名称</param>
    /// <param name="value">要插入的内容文本</param>
    /// <returns>新创建的条目封装对象</returns>
    IWordAutoTextEntry Add(string name, string value);

    /// <summary>
    /// 检查是否存在指定名称的自动图文集条目
    /// </summary>
    /// <param name="name">要查找的条目名称</param>
    /// <returns>是否存在</returns>
    bool Contains(string name);

    /// <summary>
    /// 获取所有自动图文集条目的名称列表
    /// </summary>
    /// <returns>名称列表</returns>
    List<string> GetNames();

    /// <summary>
    /// 清空所有自动图文集条目（谨慎使用）
    /// </summary>
    void Clear();
}