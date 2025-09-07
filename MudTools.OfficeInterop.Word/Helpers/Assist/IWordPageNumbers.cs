//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示文档中所有页码的对象集合。
/// <para>注：使用 Headers 或 Footers 对象的 PageNumbers 属性可返回 PageNumbers 集合。</para>
/// <para>注：使用 PageNumbers(index)（其中 index 是索引号）可返回单个 PageNumber 对象。</para>
/// </summary>
public interface IWordPageNumbers : IEnumerable<IWordPageNumber>, IDisposable
{
    #region 基本属性 (Basic Properties)

    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取集合中的页码数量。
    /// </summary>
    int Count { get; }

    #endregion

    #region 集合索引器 (Collection Indexer)

    /// <summary>
    /// 通过索引号获取单个页码。
    /// </summary>
    /// <param name="index">页码的索引号（从 1 开始）。</param>
    /// <returns>指定的页码对象。</returns>
    IWordPageNumber this[int index] { get; }

    #endregion

    #region 页码集合属性 (Page Numbers Collection Properties)
    /// <summary>
    /// 获取或设置一个值，该值指示是否在首页显示页码。
    /// </summary>
    bool ShowFirstPageNumber { get; set; }

    /// <summary>
    /// 获取或设置页码的起始编号。
    /// </summary>
    int StartingNumber { get; set; }

    #endregion

    #region 页码集合方法 (Page Numbers Collection Methods)

    /// <summary>
    /// 将新的页码添加到集合中。
    /// </summary>
    /// <param name="alignment">页码的对齐方式。</param>
    /// <param name="pageNumbers">页码对象。</param>
    /// <returns>表示添加的页码的对象。</returns>
    IWordPageNumber Add(WdPageNumberAlignment alignment, int pageNumbers);

    #endregion
}