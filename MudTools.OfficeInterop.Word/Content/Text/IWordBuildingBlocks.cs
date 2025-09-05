//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示对 Microsoft Word 中 BuildingBlocks 集合的封装接口。
/// 该集合包含某一特定类型（如页眉）和类别（如“常规”）下的所有构建基块条目。
/// </summary>
public interface IWordBuildingBlocks : IEnumerable<IWordBuildingBlock>, IDisposable
{
    /// <summary>
    /// 获取集合中构建基块的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过 1-based 索引获取指定位置的构建基块。
    /// </summary>
    /// <param name="index">索引（从1开始）。</param>
    /// <returns>封装后的构建基块对象，若不存在则返回 null。</returns>
    IWordBuildingBlock? this[int index] { get; }

    /// <summary>
    /// 通过名称获取指定的构建基块（区分大小写）。
    /// </summary>
    /// <param name="name">构建基块的名称。</param>
    /// <returns>封装后的构建基块对象，若不存在则返回 null。</returns>
    IWordBuildingBlock? this[string name] { get; }

    /// <summary>
    /// 向当前集合（即当前类型+类别）中添加一个新的构建基块。
    /// 如果指定名称已存在，Word 会抛出异常。
    /// </summary>
    /// <param name="name">构建基块的名称。</param>
    /// <param name="value">要添加的内容文本。</param>
    /// <param name="insertOptions"></param>
    /// <returns>新创建的封装构建基块对象。</returns>
    /// <exception cref="ArgumentException">当 name 为空时抛出。</exception>
    /// <exception cref="InvalidOperationException">当添加失败时抛出。</exception>
    public IWordBuildingBlock Add(string name, string value, WdDocPartInsertOptions insertOptions);

    /// <summary>
    /// 判断当前集合中是否存在指定名称的构建基块。
    /// </summary>
    /// <param name="name">要查找的名称。</param>
    /// <returns>存在返回 true，否则 false。</returns>
    bool Contains(string name);

    /// <summary>
    /// 获取当前集合中所有构建基块的名称列表。
    /// </summary>
    /// <returns>名称字符串列表。</returns>
    List<string> GetNames();

    /// <summary>
    /// 删除当前集合中的所有构建基块。
    /// </summary>
    void Clear();
}