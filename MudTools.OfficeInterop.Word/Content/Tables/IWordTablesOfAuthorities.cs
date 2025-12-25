//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中所有引文目录 (Table of Authorities, TOA) 的集合的二次封装接口。
/// 此接口允许枚举、访问特定引文目录，并向文档中添加新的引文目录 [[1]]。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordTablesOfAuthorities : IEnumerable<IWordTableOfAuthorities?>, IDisposable
{
    /// <summary>
    /// 获取此引文目录集合所属的 Word 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此引文目录集合的父对象（通常是 <see cref="IWordDocument"/>）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取文档中引文目录的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取集合中指定索引处的引文目录。索引从 1 开始。
    /// </summary>
    /// <param name="index">引文目录的索引（从 1 开始）。</param>
    /// <returns>指定索引处的 <see cref="IWordTableOfAuthorities"/> 对象，如果索引无效则返回 null。</returns>
    IWordTableOfAuthorities? this[int index] { get; }

    /// <summary>
    /// 在指定的范围内向文档添加一个新的引文目录。
    /// </summary>
    /// <param name="range">要插入引文目录的位置。如果范围未折叠，引文目录将替换该范围 [[14]]。</param>
    /// <param name="category">可选。要包含的条目类别（1-16）。默认为 0，表示包含所有类别。</param>
    /// <param name="bookmark">可选。仅包含指定书签范围内的引文。</param>
    /// <param name="entrySeparator">可选。分隔条目和页码的字符（最多五个）[[16]]。</param>
    /// <param name="pageRangeSeparator">可选。分隔页码范围的字符（最多五个）。</param>
    /// <param name="passim">可选。如果为 true，则在多页上出现的条目后显示 "passim"。</param>
    /// <returns>新创建的 <see cref="IWordTableOfAuthorities"/> 对象。</returns>
    /// <exception cref="ArgumentNullException">当 <paramref name="range"/> 为 null 时抛出。</exception>
    /// <exception cref="InvalidOperationException">当添加引文目录操作失败时抛出。</exception>
    IWordTableOfAuthorities? Add(
        IWordRange range,
        int category = 0,
        string? bookmark = null,
        string? entrySeparator = null,
        string? pageRangeSeparator = null,
        bool passim = false);

    /// <summary>
    /// 在文档中为指定的文本标记一个引文条目。
    /// 这会在文档中插入一个 TA (Table of Authorities Entry) 域 [[30]]。
    /// </summary>
    /// <param name="range">要标记为引文的文本范围。</param>
    /// <param name="entry">引文条目的长名称。</param>
    /// <param name="shortCitation">引文条目的短名称（可选）。</param>
    /// <param name="category">条目类别（1-16）（可选）。</param>
    /// <returns>插入的 TA 域对象（封装为 <see cref="IWordField"/>）。</returns>
    IWordField? MarkCitation(
        IWordRange range,
        string entry,
        string? shortCitation = null,
        int category = 0);
}