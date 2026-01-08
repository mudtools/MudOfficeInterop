//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 脚注集合的封装接口。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordFootnotes : IEnumerable<IWordFootnote?>, IOfficeObject<IWordFootnotes, MsWord.Footnotes>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取脚注数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取或设置所有脚注的位置。
    /// </summary>
    WdFootnoteLocation Location { get; set; }

    /// <summary>
    /// 获取或设置选择范围、范围或文档中脚注的编号样式。
    /// </summary>
    WdNoteNumberStyle NumberStyle { get; set; }

    /// <summary>
    /// 获取或设置起始注释编号。
    /// </summary>
    int StartingNumber { get; set; }

    /// <summary>
    /// 获取或设置在分页符或分节符后脚注的编号方式。
    /// </summary>
    WdNumberingRule NumberingRule { get; set; }

    /// <summary>
    /// 获取表示脚注分隔符的范围对象。
    /// </summary>
    IWordRange? Separator { get; }

    /// <summary>
    /// 获取表示脚注续行分隔符的范围对象。
    /// </summary>
    IWordRange? ContinuationSeparator { get; }

    /// <summary>
    /// 获取表示脚注续行通知的范围对象。
    /// </summary>
    IWordRange? ContinuationNotice { get; }

    /// <summary>
    /// 通过索引获取集合中的单个脚注对象。
    /// </summary>
    /// <param name="index">指示单个对象在集合中位置的整数索引。</param>
    /// <returns>指定索引处的脚注对象。</returns>
    IWordFootnote? this[int index] { get; }

    /// <summary>
    /// 向范围添加脚注，并返回表示该脚注的Footnote对象。
    /// </summary>
    /// <param name="range">标记为尾注或脚注的范围。可以是折叠的范围。</param>
    /// <param name="reference">自定义引用标记的文本。如果省略此参数，Microsoft Word将插入自动编号的引用标记。</param>
    /// <param name="text">尾注或脚注的文本。</param>
    /// <returns>新添加的脚注对象。</returns>
    IWordFootnote? Add(IWordRange range, string? reference = null, string? text = null);

    /// <summary>
    /// 将脚注转换为尾注。
    /// </summary>
    void Convert();

    /// <summary>
    /// 将文档中的所有脚注转换为尾注，反之亦然。
    /// </summary>
    void SwapWithEndnotes();

    /// <summary>
    /// 将脚注分隔符重置为默认分隔符。默认分隔符是一条短水平线，用于分隔文档文本和注释。
    /// </summary>
    void ResetSeparator();

    /// <summary>
    /// 将脚注续行分隔符重置为默认分隔符。默认分隔符是一条长水平线，用于分隔文档文本和从前一页继续的注释。
    /// </summary>
    void ResetContinuationSeparator();

    /// <summary>
    /// 将脚注续行通知重置为默认通知。默认通知为空白（无文本）。
    /// </summary>
    void ResetContinuationNotice();
}