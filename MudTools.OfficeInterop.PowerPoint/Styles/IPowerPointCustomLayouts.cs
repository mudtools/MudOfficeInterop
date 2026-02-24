//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 幻灯片母版中的自定义版式集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointCustomLayouts : IEnumerable<IPowerPointCustomLayout?>, IOfficeObject<IPowerPointCustomLayouts, MsPowerPoint.CustomLayouts>, IDisposable
{
    /// <summary>
    /// 获取集合的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此集合的父对象。
    /// </summary>
    /// <value>表示此页眉页脚集合父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中的自定义版式数量。
    /// </summary>
    /// <value>集合中的自定义版式数量。</value>
    int Count { get; }

    /// <summary>
    /// 通过索引或名称获取集合中的指定自定义版式。
    /// </summary>
    /// <param name="index">要获取的自定义版式的索引（从1开始）或名称。</param>
    /// <value>指定索引或名称对应的 <see cref="IPowerPointCustomLayout"/> 对象。</value>
    IPowerPointCustomLayout? this[int index] { get; }

    /// <summary>
    /// 通过索引或名称获取集合中的指定自定义版式。
    /// </summary>
    /// <param name="name">要获取的自定义版式的索引（从1开始）或名称。</param>
    /// <value>指定索引或名称对应的 <see cref="IPowerPointCustomLayout"/> 对象。</value>
    IPowerPointCustomLayout? this[string name] { get; }

    /// <summary>
    /// 在自定义版式集合中添加新自定义版式。
    /// </summary>
    /// <param name="index">新自定义版式要插入的位置索引。</param>
    /// <returns>新添加的 <see cref="IPowerPointCustomLayout"/> 对象。</returns>
    IPowerPointCustomLayout? Add(int index);

    /// <summary>
    /// 粘贴剪贴板内容为自定义版式。
    /// </summary>
    /// <param name="index">粘贴自定义版式的位置索引。值为-1表示在末尾添加。</param>
    /// <returns>粘贴的 <see cref="IPowerPointCustomLayout"/> 对象。</returns>
    IPowerPointCustomLayout? Paste(int index = -1);
}
