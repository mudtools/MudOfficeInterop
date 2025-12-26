//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 Word 尾注集合的封装接口。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordEndnotes : IEnumerable<IWordEndnote?>, IOfficeObject<IWordEndnotes>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取尾注数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取尾注。
    /// </summary>
    IWordEndnote? this[int index] { get; }

    /// <summary>
    /// 获取尾注分隔符。
    /// </summary>
    IWordRange? Separator { get; }

    /// <summary>
    /// 获取尾注续分隔符。
    /// </summary>
    IWordRange? ContinuationSeparator { get; }

    /// <summary>
    /// 获取尾注续说明。
    /// </summary>
    IWordRange? ContinuationNotice { get; }

    /// <summary>
    /// 获取或设置尾注编号方式。
    /// </summary>
    WdNoteNumberStyle NumberStyle { get; set; }

    /// <summary>
    /// 获取或设置尾注起始编号。
    /// </summary>
    int StartingNumber { get; set; }

    /// <summary>
    /// 获取或设置尾注编号格式。
    /// </summary>
    WdNumberingRule NumberingRule { get; set; }

    /// <summary>
    /// 获取或设置尾注位置。
    /// </summary>
    WdEndnoteLocation Location { get; set; }

    /// <summary>
    /// 添加新的尾注。
    /// </summary>
    /// <param name="range">添加尾注的范围。</param>
    /// <param name="referenceText">引用文本。</param>
    /// <param name="noteText">尾注文本。</param>
    /// <returns>新创建的尾注。</returns>
    IWordEndnote? Add(IWordRange range, string? referenceText = null, string? noteText = null);

    /// <summary>
    /// 将尾注转换为内嵌格式（inline）或从内嵌格式转换回来
    /// </summary>
    void Convert();

    /// <summary>
    /// 与脚注交换位置，将尾注转换为脚注或将脚注转换为尾注
    /// </summary>
    void SwapWithFootnotes();

    /// <summary>
    /// 重置尾注分隔符为默认样式
    /// </summary>
    void ResetSeparator();

    /// <summary>
    /// 重置尾注续分隔符为默认样式
    /// </summary>
    void ResetContinuationSeparator();

    /// <summary>
    /// 重置尾注续说明为默认样式
    /// </summary>
    void ResetContinuationNotice();
}