//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordXMLNodes : IOfficeObject<IWordXMLNodes>, IEnumerable<IWordXMLNode>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中XML节点的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取XML节点。
    /// </summary>
    /// <param name="index">要获取的XML节点的从零开始的索引。</param>
    /// <returns>指定索引处的XML节点；如果索引超出范围，则返回null。</returns>
    IWordXMLNode? this[int index] { get; }


    /// <summary>
    /// 向集合中添加一个新的XML节点。
    /// </summary>
    /// <param name="name">要添加的XML节点的名称。</param>
    /// <param name="namespaces">XML节点使用的命名空间。</param>
    /// <param name="range">XML节点关联的Word范围对象，可为null。</param>
    /// <returns>新添加的XML节点；如果添加失败，则返回null。</returns>
    IWordXMLNode? Add(string name, string namespaces, IWordRange? range);

}