//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中的审阅者集合，提供对文档中各个审阅者的访问和管理功能
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordReviewers : IEnumerable<IWordReviewer?>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取审阅者集合中的审阅者数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据指定的索引获取审阅者对象。
    /// </summary>
    /// <param name="index">审阅者在集合中的从零开始的索引。</param>
    /// <returns>位于指定索引处的审阅者对象，如果不存在则返回null。</returns>
    IWordReviewer? this[int index] { get; }

    /// <summary>
    /// 根据指定的名称获取审阅者对象。
    /// </summary>
    /// <param name="name">要查找的审阅者名称。</param>
    /// <returns>具有指定名称的审阅者对象，如果不存在则返回null。</returns>
    IWordReviewer? this[string name] { get; }
}