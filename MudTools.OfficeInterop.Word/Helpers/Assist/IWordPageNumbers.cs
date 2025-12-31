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
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordPageNumbers : IEnumerable<IWordPageNumber?>, IOfficeObject<IWordPageNumbers>, IDisposable
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
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取指定集合中的项目数。
    /// </summary>
    int Count { get; }

    #endregion


    /// <summary>
    /// 获取或设置 PageNumbers 对象的数字样式。
    /// </summary>
    WdPageNumberStyle NumberStyle { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示页码或题注标签是否包含章节号。
    /// </summary>
    bool IncludeChapterNumber { get; set; }

    /// <summary>
    /// 获取或设置应用于文档章节标题的标题级别样式。可以是 0 到 8 之间的数字，对应标题级别 1 到 9。
    /// </summary>
    int HeadingLevelForChapter { get; set; }

    /// <summary>
    /// 获取或设置用于分隔章节号和页码的分隔符字符。可以是 WdSeparatorType 常量之一。
    /// </summary>
    WdSeparatorType ChapterPageSeparator { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在指定节的开头是否重新从 1 开始页码编号。
    /// </summary>
    bool RestartNumberingAtSection { get; set; }

    /// <summary>
    /// 获取或设置起始注释编号、行号或页码。
    /// </summary>
    int StartingNumber { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示页码是否出现在节的第一页上。
    /// </summary>
    bool ShowFirstPageNumber { get; set; }

    /// <summary>
    /// 返回集合中的单个对象。
    /// </summary>
    /// <param name="index">指示单个对象序数位置的整数。</param>
    /// <returns>指定索引处的 PageNumber 对象。</returns>
    IWordPageNumber? this[int index] { get; }

    /// <summary>
    /// 返回表示添加到节中页眉或页脚的页码的 PageNumber 对象。
    /// </summary>
    /// <param name="pageNumberAlignment">可选项。可以是任何 WdPageNumberAlignment 常量。</param>
    /// <param name="firstPage">可选项。False 表示使首页页眉和首页页脚与文档中所有后续页的页眉和页脚不同。如果 FirstPage 设置为 False，则不向第一页添加页码。如果省略此参数，则设置由 PageSetup.DifferentFirstPageHeaderFooter 属性控制。</param>
    /// <returns>新创建的 PageNumber 对象。</returns>
    IWordPageNumber? Add(WdPageNumberAlignment? pageNumberAlignment = null, bool? firstPage = null);

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否将指定的 PageNumbers 对象用双引号（"）括起来。
    /// </summary>
    bool DoubleQuote { get; set; }
}