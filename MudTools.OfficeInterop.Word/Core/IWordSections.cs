//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示文档中所有节的对象集合。
/// <para>注：使用 Document.Sections 属性可返回 Sections 集合。</para>
/// <para>注：使用 Sections(index)（其中 index 是节的索引号）可返回单个 Section 对象。索引号代表节在文档中的位置，主文档节的索引号为 1。</para>
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordSections : IEnumerable<IWordSection?>, IOfficeObject<IWordSections>, IDisposable
{
    #region 基本属性 (Basic Properties)

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
    /// 获取集合中的节数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取或设置与该节关联的页面设置属性。
    /// </summary>
    IWordPageSetup? PageSetup { get; }

    /// <summary>
    /// 获取集合中的最后一个节。
    /// </summary>
    IWordSection? Last { get; }

    /// <summary>
    /// 获取集合中的第一个节。
    /// </summary>
    IWordSection? First { get; }

    #endregion

    #region 集合索引器 (Collection Indexer)

    /// <summary>
    /// 通过索引号获取单个节。
    /// </summary>
    /// <param name="index">节的索引号（从 1 开始）。</param>
    /// <returns>指定的节对象。</returns>
    IWordSection? this[int index] { get; }

    #endregion

    #region 节集合方法 (Sections Collection Methods)

    /// <summary>
    /// 将新的节添加到文档中。
    /// </summary>
    /// <param name="range">指定新节插入位置的范围。新节将插入到此范围之前。</param>
    /// <param name="start">新节的起始位置。</param>
    /// <returns>表示添加的节的对象。</returns>
    IWordSection? Add(IWordRange range, WdSectionStart start);
    #endregion
}