//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 段落的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordParagraph : IOfficeObject<IWordParagraph, MsWord.Paragraph>, IDisposable
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
    /// 获取或设置表示指定段落格式设置的ParagraphFormat对象。
    /// </summary>
    IWordParagraphFormat? Format { get; set; }

    /// <summary>
    /// 获取或设置表示指定段落所有自定义制表位的TabStops集合。
    /// </summary>
    IWordTabStops? TabStops { get; set; }

    /// <summary>
    /// 获取或设置表示指定对象所有边框的Borders集合。
    /// </summary>
    IWordBorders? Borders { get; set; }

    /// <summary>
    /// 获取表示指定段落首字下沉的DropCap对象。
    /// </summary>
    IWordDropCap? DropCap { get; }

    /// <summary>
    /// 获取或设置指定对象的样式。
    /// </summary>
    [ComPropertyWrap(IsMethod = true, NeedConvert = true)]
    IWordStyle? Style { get; set; }

    /// <summary>
    /// 获取或设置指定对象的样式。
    /// </summary>
    [ComPropertyWrap(IsMethod = true, PropertyName = "Style", NeedConvert = true)]
    WdBuiltinStyle StyleType { get; set; }

    /// <summary>
    /// 获取或设置指定对象的样式。
    /// </summary>
    [ComPropertyWrap(IsMethod = true, PropertyName = "Style", NeedConvert = true)]
    string? StyleName { get; set; }

    /// <summary>
    /// 获取或设置表示指定段落对齐方式的WdParagraphAlignment常量。
    /// </summary>
    WdParagraphAlignment Alignment { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示当Microsoft Word重新分页时，指定段落的所有行是否保持在同一页上。
    /// </summary>
    int KeepTogether { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示当Microsoft Word重新分页时，指定段落是否与后续段落保持在同一页上。
    /// </summary>
    int KeepWithNext { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否在指定段落前强制分页。
    /// </summary>
    int PageBreakBefore { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否取消指定段落的行号显示。
    /// </summary>
    int NoLineNumber { get; set; }

    /// <summary>
    /// 获取或设置指定段落的右缩进（以磅为单位）。
    /// </summary>
    float RightIndent { get; set; }

    /// <summary>
    /// 获取或设置指定段落的左缩进值（以磅为单位）。
    /// </summary>
    float LeftIndent { get; set; }

    /// <summary>
    /// 获取或设置首行缩进或悬挂缩进的值（以磅为单位）。
    /// </summary>
    float FirstLineIndent { get; set; }

    /// <summary>
    /// 获取或设置指定段落的行距（以磅为单位）。
    /// </summary>
    float LineSpacing { get; set; }

    /// <summary>
    /// 获取或设置指定段落的行距规则。
    /// </summary>
    WdLineSpacing LineSpacingRule { get; set; }

    /// <summary>
    /// 获取或设置指定段落前的间距（以磅为单位）。
    /// </summary>
    float SpaceBefore { get; set; }

    /// <summary>
    /// 获取或设置指定段落或文本列后的间距（以磅为单位）。
    /// </summary>
    float SpaceAfter { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示指定段落是否包含在自动连字符中。
    /// </summary>
    int Hyphenation { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示当Microsoft Word重新分页时，指定段落的首行和末行是否与段落其余部分保持在同一页上。
    /// </summary>
    int WidowControl { get; set; }

    /// <summary>
    /// 获取表示指定对象底纹格式设置的Shading对象。
    /// </summary>
    IWordShading? Shading { get; }

    /// <summary>
    /// 获取或设置一个值，指示Microsoft Word是否对指定段落应用东亚换行规则。
    /// </summary>
    int FarEastLineBreakControl { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在指定段落或文本框中，Microsoft Word是否在拉丁文本的单词中间换行。
    /// </summary>
    int WordWrap { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否为指定段落启用悬挂标点。
    /// </summary>
    int HangingPunctuation { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示对于指定段落，Microsoft Word是否将行首的标点符号更改为半角字符。
    /// </summary>
    int HalfWidthPunctuationOnTopOfLine { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示Microsoft Word是否自动为指定段落添加日语和拉丁文本之间的空格。
    /// </summary>
    int AddSpaceBetweenFarEastAndAlpha { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示Microsoft Word是否自动为指定段落添加日语文本和数字之间的空格。
    /// </summary>
    int AddSpaceBetweenFarEastAndDigit { get; set; }

    /// <summary>
    /// 获取或设置表示行上字体垂直位置的WdBaselineAlignment常量。
    /// </summary>
    WdBaselineAlignment BaseLineAlignment { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示如果指定了每行的字符数，Microsoft Word是否自动调整指定段落的右缩进。
    /// </summary>
    int AutoAdjustRightIndent { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示当指定每页的行数时，Microsoft Word是否将指定段落中的字符与行网格对齐。
    /// </summary>
    int DisableLineHeightGrid { get; set; }

    /// <summary>
    /// 获取或设置指定段落的大纲级别。
    /// </summary>
    WdOutlineLevel OutlineLevel { get; set; }

    /// <summary>
    /// 获取或设置指定段落的右缩进值（以字符为单位）。
    /// </summary>
    float CharacterUnitRightIndent { get; set; }

    /// <summary>
    /// 获取或设置指定段落的左缩进值（以字符为单位）。
    /// </summary>
    float CharacterUnitLeftIndent { get; set; }

    /// <summary>
    /// 获取或设置首行缩进或悬挂缩进的值（以字符为单位）。
    /// </summary>
    float CharacterUnitFirstLineIndent { get; set; }

    /// <summary>
    /// 获取或设置指定段落前的间距量（以网格线为单位）。
    /// </summary>
    float LineUnitBefore { get; set; }

    /// <summary>
    /// 获取或设置指定段落后的间距量（以网格线为单位）。
    /// </summary>
    float LineUnitAfter { get; set; }

    /// <summary>
    /// 获取或设置指定段落的阅读顺序，而不更改其对齐方式。
    /// </summary>
    WdReadingOrder ReadingOrder { get; set; }

    /// <summary>
    /// 获取或设置将当前文档另存为网页时指定对象的标识标签。
    /// </summary>
    string ID { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示Microsoft Word是否自动设置指定段落前的间距量。
    /// </summary>
    int SpaceBeforeAuto { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示Microsoft Word是否自动设置指定段落后的间距量。
    /// </summary>
    int SpaceAfterAuto { get; set; }

    /// <summary>
    /// 获取一个值，指示段落是否包含允许Microsoft Word似乎连接不同段落样式的特殊隐藏段落标记。
    /// </summary>
    bool IsStyleSeparator { get; }

    /// <summary>
    /// 获取或设置一个值，指示左右缩进是否等宽。可以是True、False或WdConstants.wdUndefined。
    /// </summary>
    int MirrorIndents { get; set; }

    /// <summary>
    /// 获取或设置文本围绕形状或文本框的紧密程度。
    /// </summary>
    WdTextboxTightWrap TextboxTightWrap { get; set; }

    /// <summary>
    /// 获取或设置段落的折叠状态。
    /// </summary>
    bool CollapsedState { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示默认情况下是否折叠标题。
    /// </summary>
    bool CollapseHeadingByDefault { get; set; }

    /// <summary>
    /// 移除指定段落前的所有间距。
    /// </summary>
    void CloseUp();

    /// <summary>
    /// 将指定段落前的间距设置为12磅。
    /// </summary>
    void OpenUp();

    /// <summary>
    /// 如果指定段落前的间距为0，则将其设置为12磅；如果大于0，则设置为0。
    /// </summary>
    void OpenOrCloseUp();

    /// <summary>
    /// 将悬挂缩进设置为指定数量的制表位。
    /// </summary>
    /// <param name="count">要缩进的制表位数（如果为正数）或要从缩进中移除的制表位数（如果为负数）。</param>
    void TabHangingIndent(short count);

    /// <summary>
    /// 将指定段落的左缩进设置为指定数量的制表位。
    /// </summary>
    /// <param name="count">要缩进的制表位数（如果为正数）或要从缩进中移除的制表位数（如果为负数）。</param>
    void TabIndent(short count);

    /// <summary>
    /// 移除手动段落格式设置（未使用样式应用的格式设置）。
    /// </summary>
    void Reset();

    /// <summary>
    /// 将指定段落设置为单倍行距。确切行距由每段中最大字符的字体大小决定。
    /// </summary>
    void Space1();

    /// <summary>
    /// 将指定段落格式化为1.5倍行距。
    /// </summary>
    void Space15();

    /// <summary>
    /// 将指定段落设置为双倍行距。
    /// </summary>
    void Space2();

    /// <summary>
    /// 按指定字符数缩进一个或多个段落。
    /// </summary>
    /// <param name="count">要缩进的字符数。</param>
    void IndentCharWidth(short count);

    /// <summary>
    /// 按指定字符数缩进一个或多个段落的首行。
    /// </summary>
    /// <param name="count">要缩进的字符数。</param>
    void IndentFirstLineCharWidth(short count);

    /// <summary>
    /// 返回表示下一个段落的Paragraph对象。
    /// </summary>
    /// <param name="count">要向前移动的段落数。默认值为1。</param>
    /// <returns>下一个段落对象。</returns>
    IWordParagraph? Next(int? count = null);

    /// <summary>
    /// 将前一个段落作为Paragraph对象返回。
    /// </summary>
    /// <param name="count">要向后移动的段落数。默认值为1。</param>
    /// <returns>前一个段落对象。</returns>
    IWordParagraph? Previous(int? count = null);

    /// <summary>
    /// 将前一个标题级别样式（标题1到标题8）应用于指定段落。
    /// </summary>
    void OutlinePromote();

    /// <summary>
    /// 将下一个标题级别样式（标题1到标题8）应用于指定段落。
    /// </summary>
    void OutlineDemote();

    /// <summary>
    /// 通过应用普通样式将指定段落降级为正文文本。
    /// </summary>
    void OutlineDemoteToBody();

    /// <summary>
    /// 将一个或多个段落缩进一个级别。
    /// </summary>
    void Indent();

    /// <summary>
    /// 移除一个或多个段落的缩进级别。
    /// </summary>
    void Outdent();

    /// <summary>
    /// 选择列表中的编号或项目符号。
    /// </summary>
    void SelectNumber();

    /// <summary>
    /// 设置列表中段落的列表级别。
    /// </summary>
    /// <param name="level1">第一级列表级别。</param>
    /// <param name="level2">第二级列表级别。</param>
    /// <param name="level3">第三级列表级别。</param>
    /// <param name="level4">第四级列表级别。</param>
    /// <param name="level5">第五级列表级别。</param>
    /// <param name="level6">第六级列表级别。</param>
    /// <param name="level7">第七级列表级别。</param>
    /// <param name="level8">第八级列表级别。</param>
    /// <param name="level9">第九级列表级别。</param>
    void ListAdvanceTo(short level1 = 0, short level2 = 0, short level3 = 0, short level4 = 0, short level5 = 0, short level6 = 0, short level7 = 0, short level8 = 0, short level9 = 0);

    /// <summary>
    /// 将使用自定义列表级别的段落重置为原始级别设置。
    /// </summary>
    void ResetAdvanceTo();

    /// <summary>
    /// 将列表分为两个独立的列表。对于编号列表，新列表将重新从起始编号（通常为1）开始编号。
    /// </summary>
    void SeparateList();

    /// <summary>
    /// 将列表段落与指定段落上方或下方最接近的列表合并。
    /// </summary>
    void JoinList();
}