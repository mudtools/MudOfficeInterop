
namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示用于构建目录或图表目录的标题样式集合的二次封装接口。
/// 此集合包含除内置“标题 1-9”之外的、用于编译目录的其他样式 [[1]]。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordHeadingStyles : IEnumerable<IWordHeadingStyle?>, IOfficeObject<IWordHeadingStyles, MsWord.HeadingStyles>, IDisposable
{
    /// <summary>
    /// 获取此标题样式集合所属的 Word 应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此标题样式集合的父对象（通常是 <see cref="IWordTableOfContents"/>）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中标题样式的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取集合中指定索引处的标题样式。索引从 1 开始。
    /// </summary>
    /// <param name="index">标题样式的索引（从 1 开始）。</param>
    /// <returns>指定索引处的 <see cref="IWordHeadingStyle"/> 对象，如果索引无效则返回 null。</returns>
    IWordHeadingStyle? this[int index] { get; }

    /// <summary>
    /// 向集合中添加一个新的标题样式。
    /// 新添加的标题样式将在编译目录或图表目录时被包含 [[8]]。
    /// </summary>
    /// <param name="style">要添加的样式。可以是样式名称（字符串）、WdBuiltinStyle 枚举值或 <see cref="IWordStyle"/> 对象。</param>
    /// <param name="level">该样式在目录中的层级（1-9）。</param>
    /// <returns>新创建的 <see cref="IWordHeadingStyle"/> 对象。</returns>
    /// <exception cref="ArgumentNullException">当 <paramref name="style"/> 为 null 时抛出。</exception>
    /// <exception cref="ArgumentOutOfRangeException">当 <paramref name="level"/> 不在 1-9 范围内时抛出。</exception>
    /// <exception cref="InvalidOperationException">当添加操作失败时抛出。</exception>
    IWordHeadingStyle? Add(IWordStyle style, short level);

    /// <summary>
    /// 向集合中添加一个新的标题样式。
    /// 新添加的标题样式将在编译目录或图表目录时被包含 [[8]]。
    /// </summary>
    /// <param name="styleName">要添加的样式。可以是样式名称（字符串）、WdBuiltinStyle 枚举值或 <see cref="IWordStyle"/> 对象。</param>
    /// <param name="level">该样式在目录中的层级（1-9）。</param>
    /// <returns>新创建的 <see cref="IWordHeadingStyle"/> 对象。</returns>
    /// <exception cref="ArgumentNullException">当 <paramref name="styleName"/> 为 null 时抛出。</exception>
    /// <exception cref="ArgumentOutOfRangeException">当 <paramref name="level"/> 不在 1-9 范围内时抛出。</exception>
    /// <exception cref="InvalidOperationException">当添加操作失败时抛出。</exception>
    IWordHeadingStyle? Add(string styleName, short level);

    /// <summary>
    /// 向集合中添加一个新的标题样式。
    /// 新添加的标题样式将在编译目录或图表目录时被包含 [[8]]。
    /// </summary>
    /// <param name="styleName">要添加的样式。可以是样式名称（字符串）、WdBuiltinStyle 枚举值或 <see cref="IWordStyle"/> 对象。</param>
    /// <param name="level">该样式在目录中的层级（1-9）。</param>
    /// <returns>新创建的 <see cref="IWordHeadingStyle"/> 对象。</returns>
    /// <exception cref="ArgumentNullException">当 <paramref name="styleName"/> 为 null 时抛出。</exception>
    /// <exception cref="ArgumentOutOfRangeException">当 <paramref name="level"/> 不在 1-9 范围内时抛出。</exception>
    /// <exception cref="InvalidOperationException">当添加操作失败时抛出。</exception>
    IWordHeadingStyle? Add(WdBuiltinStyle styleName, short level);
}