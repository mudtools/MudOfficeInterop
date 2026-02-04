//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 演示文稿中的命名幻灯片放映集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointNamedSlideShows : IOfficeObject<IPowerPointNamedSlideShows, MsPowerPoint.NamedSlideShows>, IEnumerable<IPowerPointNamedSlideShow?>, IDisposable
{
    /// <summary>
    /// 获取集合中的命名幻灯片放映数量。
    /// </summary>
    /// <value>集合中的命名幻灯片放映数量。</value>
    int Count { get; }

    /// <summary>
    /// 获取创建此命名幻灯片放映集合的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此命名幻灯片放映集合的父对象。
    /// </summary>
    /// <value>表示此命名幻灯片放映集合父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 通过索引或名称获取集合中的指定命名幻灯片放映。
    /// </summary>
    /// <param name="index">要获取的命名幻灯片放映的索引（从1开始）或名称。</param>
    /// <value>指定索引或名称对应的 <see cref="IPowerPointNamedSlideShow"/> 对象。</value>
    IPowerPointNamedSlideShow? this[object index] { get; }

    /// <summary>
    /// 通过索引或名称获取集合中的指定命名幻灯片放映。
    /// </summary>
    /// <param name="name">要获取的命名幻灯片放映的索引（从1开始）或名称。</param>
    /// <value>指定索引或名称对应的 <see cref="IPowerPointNamedSlideShow"/> 对象。</value>
    IPowerPointNamedSlideShow? this[string name] { get; }

    /// <summary>
    /// 在命名幻灯片放映集合中添加新命名幻灯片放映。
    /// </summary>
    /// <param name="name">新命名幻灯片放映的名称。</param>
    /// <param name="safeArrayOfSlideIDs">包含幻灯片标识符的数组。</param>
    /// <returns>新添加的 <see cref="IPowerPointNamedSlideShow"/> 对象。</returns>
    IPowerPointNamedSlideShow? Add(string name, object safeArrayOfSlideIDs);
}