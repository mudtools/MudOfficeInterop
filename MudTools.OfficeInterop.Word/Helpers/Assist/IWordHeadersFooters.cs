//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示文档中所有页眉和页脚的对象集合。
/// <para>注：使用 Section.Headers 或 Section.Footers 属性可返回 HeadersFooters 集合。</para>
/// <para>注：使用 HeadersFooters(index)（其中 index 是页眉或页脚的索引号）可返回单个 HeaderFooter 对象。</para>
/// <para>注：索引号 1 代表首页页眉或页脚，索引号 2 代表奇数页页眉或页脚，索引号 3 代表偶数页页眉或页脚。</para>
/// </summary>
public interface IWordHeadersFooters : IEnumerable<IWordHeaderFooter>, IDisposable
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
    /// 获取集合中的页眉和页脚数量。
    /// </summary>
    int Count { get; }

    #endregion

    #region 集合索引器 (Collection Indexer)

    /// <summary>
    /// 通过索引号（1=首页, 2=奇数页, 3=偶数页）获取单个页眉或页脚。
    /// </summary>
    /// <param name="index">页眉或页脚的索引号（<see cref="MsWord.WdHeaderFooterIndex"/> 常量）。</param>
    /// <returns>指定的页眉或页脚对象。</returns>
    IWordHeaderFooter this[WdHeaderFooterIndex index] { get; }

    #endregion
}