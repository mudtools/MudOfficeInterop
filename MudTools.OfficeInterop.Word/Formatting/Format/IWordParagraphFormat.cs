//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.ParagraphFormat 的接口，用于操作 Word 文档中段落的格式设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordParagraphFormat : IOfficeObject<IWordParagraphFormat, MsWord.ParagraphFormat>, IDisposable
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
    /// 返回表示指定段落段落格式的只读 ParagraphFormat 对象。
    /// </summary>
    IWordParagraphFormat? Duplicate { get; }

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
    /// 获取或设置表示指定段落的对齐方式的 WdParagraphAlignment 常量。
    /// </summary>
    WdParagraphAlignment Alignment { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示当 Microsoft Word 重新分页文档时，指定段落中的所有行是否保持在同一页上。可以是 True、False 或 wdUndefined。
    /// </summary>
    int KeepTogether { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示当 Microsoft Word 重新分页文档时，指定段落是否与后面的段落保持在同一页上。可以是 True、False 或 wdUndefined。
    /// </summary>
    int KeepWithNext { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否在指定段落前强制分页。可以是 True、False 或 wdUndefined。
    /// </summary>
    int PageBreakBefore { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否为指定段落禁止行号。可以是 True、False 或 wdUndefined。
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
    /// 获取或设置首行缩进或悬挂缩进的值（以磅为单位）。使用正值设置首行缩进，使用负值设置悬挂缩进。
    /// </summary>
    float FirstLineIndent { get; set; }

    /// <summary>
    /// 获取或设置指定段落的行距（以磅为单位）。
    /// </summary>
    float LineSpacing { get; set; }

    /// <summary>
    /// 获取或设置指定段落的行距。
    /// </summary>
    WdLineSpacing LineSpacingRule { get; set; }

    /// <summary>
    /// 获取或设置指定段落之前的间距（以磅为单位）。
    /// </summary>
    float SpaceBefore { get; set; }

    /// <summary>
    /// 获取或设置指定段落之后的间距量（以磅为单位）。
    /// </summary>
    float SpaceAfter { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示指定段落是否包含在自动断字中。False 表示指定段落要从自动断字中排除。可以是 True、False 或 wdUndefined。
    /// </summary>
    int Hyphenation { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示当 Word 重新分页文档时，指定段落的第一个和最后一行是否与段落的其余部分保持在同一页上。可以是 True、False 或 wdUndefined。
    /// </summary>
    int WidowControl { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否将东亚换行规则应用于指定段落。
    /// </summary>
    int FarEastLineBreakControl { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否在指定段落或文本框中的单词中间换行拉丁文本。
    /// </summary>
    int WordWrap { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否为指定段落启用悬挂标点。
    /// </summary>
    int HangingPunctuation { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否将行首的标点符号更改为指定段落的半角字符。
    /// </summary>
    int HalfWidthPunctuationOnTopOfLine { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否设置为自动为指定段落添加日语和拉丁文本之间的空格。
    /// </summary>
    int AddSpaceBetweenFarEastAndAlpha { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否设置为自动为指定段落添加日语文本和数字之间的空格。
    /// </summary>
    int AddSpaceBetweenFarEastAndDigit { get; set; }

    /// <summary>
    /// 获取或设置表示一行上字体垂直位置的 WdBaselineAlignment 常量。
    /// </summary>
    WdBaselineAlignment BaseLineAlignment { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示如果已指定每行字符数，Microsoft Word 是否设置为自动调整指定段落的右缩进。
    /// </summary>
    int AutoAdjustRightIndent { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示当指定每页行数时，Microsoft Word 是否将指定段落中的字符与行网格对齐。
    /// </summary>
    int DisableLineHeightGrid { get; set; }

    /// <summary>
    /// 获取或设置指定段落的大纲级别。
    /// </summary>
    WdOutlineLevel OutlineLevel { get; set; }

    /// <summary>
    /// 获取或设置指定段落的阅读顺序而不更改其对齐方式。
    /// </summary>
    WdReadingOrder ReadingOrder { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否自动设置指定段落之前的间距量。
    /// </summary>
    int SpaceBeforeAuto { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否自动设置指定段落之后的间距量。
    /// </summary>
    int SpaceAfterAuto { get; set; }

    /// <summary>
    /// 获取或设置一个整数值，指示左缩进和右缩进是否具有相同的宽度。可以是 True、False 或 wdUndefined。
    /// </summary>
    int MirrorIndents { get; set; }

    /// <summary>
    /// 获取或设置表示文本如何紧密环绕形状或文本框的 WdTextboxTightWrap 枚举值。
    /// </summary>
    WdTextboxTightWrap TextboxTightWrap { get; set; }

    /// <summary>
    /// 获取或设置指定段落是否默认折叠。
    /// </summary>
    int CollapsedByDefault { get; set; }

    /// <summary>
    /// 获取或设置指定段落的右缩进值（以字符为单位）。
    /// </summary>
    float CharacterUnitRightIndent { get; set; }

    /// <summary>
    /// 获取或设置指定段落的左缩进值（以字符为单位）。
    /// </summary>
    float CharacterUnitLeftIndent { get; set; }

    /// <summary>
    /// 获取或设置首行或悬挂缩进的值（以字符为单位）。使用正值设置首行缩进，使用负值设置悬挂缩进。
    /// </summary>
    float CharacterUnitFirstLineIndent { get; set; }

    /// <summary>
    /// 获取或设置指定段落之前的间距量（以网格线为单位）。
    /// </summary>
    float LineUnitBefore { get; set; }

    /// <summary>
    /// 获取或设置指定段落之后的间距量（以网格线为单位）。
    /// </summary>
    float LineUnitAfter { get; set; }

    /// <summary>
    /// 获取或设置表示指定段落的所有自定义制表位的 TabStops 集合。
    /// </summary>
    IWordTabStops? TabStops { get; set; }

    /// <summary>
    /// 获取或设置表示指定对象的所有边框的 Borders 集合。
    /// </summary>
    IWordBorders? Borders { get; set; }

    /// <summary>
    /// 获取引用指定对象的阴影格式的 Shading 对象。
    /// </summary>
    IWordShading? Shading { get; }

    /// <summary>
    /// 删除指定段落之前的任何间距。
    /// </summary>
    void CloseUp();

    /// <summary>
    /// 将指定段落之前的间距设置为 12 磅。
    /// </summary>
    void OpenUp();

    /// <summary>
    /// 如果指定段落之前的间距为 0（零），此方法将间距设置为 12 磅。如果段落之前的间距大于 0（零），此方法将间距设置为 0（零）。
    /// </summary>
    void OpenOrCloseUp();

    /// <summary>
    /// 将悬挂缩进设置为指定的制表位数量。
    /// </summary>
    /// <param name="count">必需 Short。要缩进的制表位数量（如果为正数）或要从缩进中删除的制表位数量（如果为负数）。</param>
    void TabHangingIndent(short count);

    /// <summary>
    /// 将指定段落的左缩进设置为指定的制表位数量。
    /// </summary>
    /// <param name="count">必需 Short。要缩进的制表位数量（如果为正数）或要从缩进中删除的制表位数量（如果为负数）。</param>
    void TabIndent(short count);

    /// <summary>
    /// 删除手动段落格式（未使用样式应用的格式）。例如，如果手动右对齐段落而基础样式具有不同的对齐方式，Reset 方法会将对齐方式更改为匹配基础样式的格式。
    /// </summary>
    void Reset();

    /// <summary>
    /// 将指定段落设置为单倍行距。确切的间距由每个段落中最大字符的字体大小决定。
    /// </summary>
    void Space1();

    /// <summary>
    /// 将指定段落格式化为 1.5 倍行距。确切的间距是通过向每个段落中最大字符的字体大小添加 6 磅来确定的。
    /// </summary>
    void Space15();

    /// <summary>
    /// 将指定段落设置为双倍行距。确切的间距是通过向每个段落中最大字符的字体大小添加 12 磅来确定的。
    /// </summary>
    void Space2();

    /// <summary>
    /// 按指定的字符数缩进一个或多个段落。
    /// </summary>
    /// <param name="count">必需 Short。要缩进指定段落的字符数。</param>
    void IndentCharWidth(short count);

    /// <summary>
    /// 按指定的字符数缩进一个或多个段落的首行。
    /// </summary>
    /// <param name="count">必需 Short。要缩进每个指定段落首行的字符数。</param>
    void IndentFirstLineCharWidth(short count);
}