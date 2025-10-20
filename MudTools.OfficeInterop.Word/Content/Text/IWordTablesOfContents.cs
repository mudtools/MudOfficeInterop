
namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中所有目录 (Table of Contents, TOC) 的集合的二次封装接口。
/// 此接口允许枚举、访问特定目录，并向文档中添加新的目录。
/// </summary>
public interface IWordTablesOfContents : IEnumerable<IWordTableOfContents>, IDisposable
{
    /// <summary>
    /// 获取此目录集合所属的 Word 应用程序对象。
    /// </summary>
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
    /// 在指定的范围内向文档添加一个新的目录。
    /// </summary>
    /// <param name="range">要插入目录的位置。如果范围未折叠，目录将替换该范围。</param>
    /// <param name="useHeadingStyles">可选。如果为 true，则使用内置标题样式来构建目录。</param>
    /// <param name="upperHeadingLevel">可选。目录中包含的最高标题级别（1-9）。</param>
    /// <param name="lowerHeadingLevel">可选。目录中包含的最低标题级别（1-9）。</param>
    /// <param name="useFields">可选。如果为 true，则目录由 TC (Table of Contents) 域构成。</param>
    /// <param name="tableId">可选。用于标识由 TC 域构成的目录的标识符。</param>
    /// <returns>新创建的 <see cref="IWordTableOfContents"/> 对象。</returns>
    /// <exception cref="ArgumentNullException">当 <paramref name="range"/> 为 null 时抛出。</exception>
    /// <exception cref="InvalidOperationException">当添加目录操作失败时抛出。</exception>
    IWordTableOfContents Add(
        IWordRange range,
        bool useHeadingStyles = true,
        int upperHeadingLevel = 1,
        int lowerHeadingLevel = 3,
        bool useFields = false,
        string? tableId = null);
}