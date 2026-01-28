//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示图示中节点的集合，提供对多个图示节点的访问和操作。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointDiagramNodes : IEnumerable<IPowerPointDiagramNode?>, IDisposable
{
    /// <summary>
    /// 获取创建此图示节点集合的应用程序。
    /// </summary>
    /// <value>表示应用程序的对象。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取创建此对象的应用程序的创建者代码。
    /// </summary>
    /// <value>创建者标识符。</value>
    int Creator { get; }

    /// <summary>
    /// 通过索引获取集合中的图示节点。
    /// </summary>
    /// <param name="index">要获取的节点的索引或名称。</param>
    /// <returns>指定索引处的图示节点。</returns>
    IPowerPointDiagramNode? this[int index] { get; }

    /// <summary>
    /// 通过索引获取集合中的图示节点。
    /// </summary>
    /// <param name="name">要获取的节点的索引或名称。</param>
    /// <returns>指定索引处的图示节点。</returns>
    IPowerPointDiagramNode? this[string name] { get; }

    /// <summary>
    /// 选择集合中的所有图示节点。
    /// </summary>
    void SelectAll();

    /// <summary>
    /// 获取图示节点集合的父对象。
    /// </summary>
    /// <value>父对象，通常是图示或图示节点。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中图示节点的数量。
    /// </summary>
    /// <value>集合中节点的总数。</value>
    int Count { get; }
}