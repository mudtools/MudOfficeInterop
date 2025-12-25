//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 提供对协同编写对象模型的主要入口点。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordCoAuthoring : IDisposable
{
    /// <summary>
    /// 获取与此协作者关联的Word应用程序实例
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此协作者的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取表示当前正在编辑文档的所有共同作者的 CoAuthors 集合。
    /// </summary>
    /// <returns>表示当前正在编辑文档的所有共同作者的 CoAuthors 集合。</returns>
    IWordCoAuthors Authors { get; }

    /// <summary>
    /// 获取表示当前用户的 CoAuthor 对象。
    /// </summary>
    /// <returns>表示当前用户的 CoAuthor 对象。</returns>
    IWordCoAuthor? Me { get; }

    /// <summary>
    /// 获取文档是否有未接受的待处理更新。
    /// </summary>
    /// <returns>如果文档有待处理但未接受的更新，则为 true；否则为 false。</returns>
    bool PendingUpdates { get; }

    /// <summary>
    /// 获取表示文档中锁定的 CoAuthLocks 集合。
    /// </summary>
    /// <returns>表示文档中锁定的 CoAuthLocks 集合。</returns>
    IWordCoAuthLocks? Locks { get; }

    /// <summary>
    /// 获取表示文档可用的所有更新的 CoAuthUpdates 集合。
    /// </summary>
    /// <returns>表示文档可用的所有更新的 CoAuthUpdates 集合。</returns>
    IWordCoAuthUpdates? Updates { get; }

    /// <summary>
    /// 获取表示文档中所有冲突的 Conflicts 集合。
    /// </summary>
    /// <returns>表示文档中所有冲突的 Conflicts 集合。</returns>
    IWordConflicts? Conflicts { get; }

    /// <summary>
    /// 获取此文档是否可以协同编写。
    /// </summary>
    /// <returns>如果此文档可以协同编写，则为 true；否则为 false。</returns>
    bool CanShare { get; }

    /// <summary>
    /// 获取文档是否可以自动合并。
    /// </summary>
    /// <returns>如果文档可以自动合并，则为 true；否则为 false。</returns>
    bool CanMerge { get; }
}