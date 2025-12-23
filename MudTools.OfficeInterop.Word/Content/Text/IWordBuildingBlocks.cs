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
[ComCollectionWrap(ComNamespace = "MsWord"), ItemIndex, NoneEnumerable]
public interface IWordBuildingBlocks : IEnumerable<IWordBuildingBlock?>, IDisposable
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
    /// 在集合中添加一个新的构建基块。
    /// </summary>
    /// <param name="name">构建基块的名称。</param>
    /// <param name="range">包含构建基块内容的Word范围对象。</param>
    /// <param name="description">构建基块的描述信息，可选参数，默认为null。</param>
    /// <param name="insertOptions">插入选项，指定如何插入构建基块，默认为wdInsertContent。</param>
    /// <returns>如果成功添加则返回封装后的构建基块对象，否则返回null。</returns>
    IWordBuildingBlock? Add(string name, IWordRange range, string? description = null, WdDocPartInsertOptions insertOptions = WdDocPartInsertOptions.wdInsertContent);

}