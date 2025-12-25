
namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中所有目录 (Table of Contents, TOC) 的集合的二次封装接口。
/// 此接口允许枚举、访问特定目录，并向文档中添加新的目录。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordTablesOfContents : IEnumerable<IWordTableOfContents?>, IDisposable
{
    /// <summary>
    /// 获取此目录集合所属的 Word 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此目录集合的父对象（通常是 <see cref="IWordDocument"/>）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取文档中目录的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取集合中指定索引处的目录。索引从 1 开始。
    /// </summary>
    /// <param name="index">目录的索引（从 1 开始）。</param>
    /// <returns>指定索引处的 <see cref="IWordTableOfContents"/> 对象，如果索引无效则返回 null。</returns>
    IWordTableOfContents? this[int index] { get; }

    /// <summary>
    /// 在指定范围后插入一个 TC（目录条目）字段。方法返回表示 TC 字段的 Field 对象。
    /// </summary>
    /// <param name="range">必需 Range 对象。条目的位置。TC 字段在 Range 之后插入。</param>
    /// <param name="entry">可选 Object。出现在目录或图表目录中的文本。要表示子条目，请包含主条目文本和子条目文本，用冒号 (:) 分隔（例如，"Introduction:The Product"）。</param>
    /// <param name="entryAutoText">可选 Object。包含索引、图表目录或目录文本的自动图文集条目名称（忽略 Entry 参数）。</param>
    /// <param name="tableID">可选 Object。图表目录或目录项目的一个字母标识符（例如，"i" 表示"插图"）。</param>
    /// <param name="level">可选 Object。目录或图表目录中条目的级别。</param>
    /// <returns>Microsoft.Office.Interop.Word.Field</returns>
    IWordField? MarkEntry(IWordRange range, string? entry = null, string? entryAutoText = null,
                          string? tableID = null, int? level = null);

    /// <summary>
    /// 返回表示添加到文档的目录的 TableOfContents 对象。
    /// </summary>
    /// <param name="range">必需 Range 对象。希望目录出现的位置范围。如果范围未折叠，目录将替换该范围。</param>
    /// <param name="useHeadingStyles">可选 Object。如果为 True，则使用内置标题样式创建目录。默认值为 True。</param>
    /// <param name="upperHeadingLevel">可选 Object。目录的起始标题级别。对应于目录（TOC）字段的 \o 开关使用的起始值。默认值为 1。</param>
    /// <param name="lowerHeadingLevel">可选 Object。目录的结束标题级别。对应于目录（TOC）字段的 \o 开关使用的结束值。默认值为 9。</param>
    /// <param name="useFields">可选 Object。如果为 True，则使用目录条目（TC）字段创建目录。使用 TablesOfContents.MarkEntry 方法标记要包含在目录中的条目。默认值为 False。</param>
    /// <param name="tableID">可选 Object。用于从 TC 字段构建目录的一个字母标识符。对应于目录（TOC）字段的 \f 开关。例如，"T" 使用表标识符 T 从 TC 字段构建目录。如果省略此参数，则不使用 TC 字段。</param>
    /// <param name="rightAlignPageNumbers">可选 Object。如果为 True，则目录中的页码右对齐。默认值为 True。</param>
    /// <param name="includePageNumbers">可选 Object。如果为 True，则在目录中包含页码。默认值为 True。</param>
    /// <param name="addedStyles">可选 Object。用于编译目录的附加样式字符串名称（Heading 1 – Heading 9 样式之外的样式）。使用 HeadingStyles 对象的 Add 方法创建新的标题样式。</param>
    /// <param name="useHyperlinks">可选 Object。如果将文档发布到 Web，目录条目应格式化为超链接，则为 True。默认值为 True。</param>
    /// <param name="hidePageNumbersInWeb">可选 Object。如果为 True，则使用大纲级别创建目录。默认值为 False。</param>
    /// <param name="useOutlineLevels">可选 Object。如果为 True，则使用大纲级别创建目录。默认值为 False。</param>
    /// <returns>Microsoft.Office.Interop.Word.TableOfContents</returns>
    IWordTableOfContents? Add(IWordRange range, bool? useHeadingStyles = null,
     int? upperHeadingLevel = null, int? lowerHeadingLevel = null,
     bool? useFields = null, string? tableID = null, bool? rightAlignPageNumbers = null,
     bool? includePageNumbers = null, string? addedStyles = null, bool? useHyperlinks = null,
     bool? hidePageNumbersInWeb = null, bool? useOutlineLevels = null);
}