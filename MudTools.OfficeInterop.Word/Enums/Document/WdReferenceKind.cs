//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定交叉引用中包含的信息
/// </summary>
[Guid("394033AF-E0BA-30E7-B099-A79873E55634")]
public enum WdReferenceKind
{
    /// <summary>
    /// 插入指定项的文本值。例如，插入指定标题的文本。
    /// </summary>
    wdContentText = -1,

    /// <summary>
    /// 插入标题或段落，并在大纲编号列表中包含足够的相对位置信息以标识该项。
    /// </summary>
    wdNumberRelativeContext = -2,

    /// <summary>
    /// 插入标题或段落，但不包含其在大纲编号列表中的相对位置。
    /// </summary>
    wdNumberNoContext = -3,

    /// <summary>
    /// 插入完整的标题或段落编号。
    /// </summary>
    wdNumberFullContext = -4,

    /// <summary>
    /// 插入指定公式、图或表的标签、编号以及任何附加题注。
    /// </summary>
    wdEntireCaption = 2,

    /// <summary>
    /// 仅插入指定公式、图或表的标签和编号。
    /// </summary>
    wdOnlyLabelAndNumber = 3,

    /// <summary>
    /// 仅插入指定公式、图或表的题注文本。
    /// </summary>
    wdOnlyCaptionText = 4,

    /// <summary>
    /// 插入脚注引用标记。
    /// </summary>
    wdFootnoteNumber = 5,

    /// <summary>
    /// 插入尾注引用标记。
    /// </summary>
    wdEndnoteNumber = 6,

    /// <summary>
    /// 插入指定项的页码。
    /// </summary>
    wdPageNumber = 7,

    /// <summary>
    /// 根据需要插入单词"Above"（上方）或"Below"（下方）。
    /// </summary>
    wdPosition = 15,

    /// <summary>
    /// 插入带格式的脚注引用标记。
    /// </summary>
    wdFootnoteNumberFormatted = 16,

    /// <summary>
    /// 插入带格式的尾注引用标记。
    /// </summary>
    wdEndnoteNumberFormatted = 17
}