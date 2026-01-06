//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中的一个连续区域。
/// <para>每个 Range 对象都由起始字符和结束字符位置定义。Range 对象独立于所选内容，可以定义多个区域。</para>
/// <para>注：此接口封装了 Microsoft.Office.Interop.Word.Range 的主要属性和方法。</para>
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordRange : IOfficeObject<IWordRange, MsWord.Range>, IDisposable
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
    /// 获取或设置区域的起始字符位置。
    /// </summary>
    int Start { get; set; }

    /// <summary>
    /// 获取或设置区域的结束字符位置。
    /// </summary>
    int End { get; set; }

    /// <summary>
    /// 获取或设置区域中的文本。
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取区域的副本。
    /// </summary>
    IWordRange? Duplicate { get; }

    /// <summary>
    /// 获取区域所属的文档。
    /// </summary>
    IWordDocument? Document { get; }

    /// <summary>
    /// 获取或设置指定对象的样式。
    /// </summary>
    [ComPropertyWrap(IsMethod = true, NeedConvert = true)]
    IWordStyle? Style { get; set; }

    /// <summary>
    /// 获取或设置指定对象的样式。
    /// </summary>
    [ComPropertyWrap(IsMethod = true, PropertyName = "Style", NeedConvert = true)]
    WdBuiltinStyle? StyleType { get; set; }

    /// <summary>
    /// 获取或设置指定对象的样式。
    /// </summary>
    [ComPropertyWrap(IsMethod = true, PropertyName = "Style", NeedConvert = true)]
    string? StyleName { get; set; }

    /// <summary>
    /// 获取区域的文章类型。
    /// </summary>
    WdStoryType? StoryType { get; }

    /// <summary>
    /// 获取区域的文章长度。
    /// </summary>
    int StoryLength { get; }

    /// <summary>
    /// 获取区域中的下一文章范围。
    /// </summary>
    IWordRange? NextStoryRange { get; }

    /// <summary>
    /// 获取区域中的单元格集合。
    /// </summary>
    IWordCells? Cells { get; }

    /// <summary>
    /// 获取区域中的节集合。
    /// </summary>
    IWordSections? Sections { get; }

    /// <summary>
    /// 获取文本检索模式对象，用于控制文本的检索方式。
    /// </summary>
    IWordTextRetrievalMode? TextRetrievalMode { get; }

    /// <summary>
    /// 获取区域中的框架集合。
    /// </summary>
    IWordFrames? Frames { get; }

    /// <summary>
    /// 获取与区域关联的同义词信息对象。
    /// </summary>
    IWordSynonymInfo? SynonymInfo { get; }

    /// <summary>
    /// 获取区域中的列集合。
    /// </summary>
    IWordColumns? Columns { get; }

    /// <summary>
    /// 获取区域中的行集合。
    /// </summary>
    IWordRows? Rows { get; }

    /// <summary>
    /// 获取一个值，指示是否可以编辑此区域。
    /// </summary>
    int CanEdit { get; }

    /// <summary>
    /// 获取一个值，指示是否可以在此区域粘贴内容。
    /// </summary>
    int CanPaste { get; }

    /// <summary>
    /// 获取一个值，指示此区域是否位于行尾标记处。
    /// </summary>
    bool IsEndOfRowMark { get; }

    /// <summary>
    /// 获取区域的书签ID（如果有）。
    /// </summary>
    int BookmarkID { get; }

    /// <summary>
    /// 获取前一个书签的ID。
    /// </summary>
    int PreviousBookmarkID { get; }


    //WdInformation Information { get; }

    #endregion

    #region 格式化属性 (Formatting Properties - 字体和段落)

    /// <summary>
    /// 获取或设置区域的字体格式。
    /// </summary>
    IWordFont? Font { get; }

    /// <summary>
    /// 获取或设置区域的段落格式。
    /// </summary>
    IWordParagraphFormat? ParagraphFormat { get; }

    /// <summary>
    /// 获取或设置区域的粗体格式 (0=False, 1=True, 9999999=Undefined)。
    /// </summary>
    int Bold { get; set; }

    /// <summary>
    /// 获取或设置区域的斜体格式 (0=False, 1=True, 9999999=Undefined)。
    /// </summary>
    int Italic { get; set; }

    /// <summary>
    /// 获取或设置区域的下划线格式。
    /// </summary>
    WdUnderline Underline { get; set; }

    /// <summary>
    /// 获取或设置区域的突出显示颜色。
    /// </summary>
    WdColorIndex HighlightColorIndex { get; set; }

    /// <summary>
    /// 获取或设置区域的字符大小写。
    /// </summary>
    WdCharacterCase Case { get; set; }

    /// <summary>
    /// 获取区域的底纹格式。
    /// </summary>
    IWordShading? Shading { get; }

    /// <summary>
    /// 获取区域中的列表格式。
    /// </summary>
    IWordListFormat? ListFormat { get; }

    /// <summary>
    /// 获取或设置区域的页面设置。
    /// </summary>
    IWordPageSetup? PageSetup { get; set; }

    #endregion

    #region 集合属性 (Collection Properties - 第一部分)

    /// <summary>
    /// 获取区域中的所有段落。
    /// </summary>
    IWordParagraphs? Paragraphs { get; }

    /// <summary>
    /// 获取区域中的所有句子。
    /// </summary>
    IWordSentences? Sentences { get; }

    /// <summary>
    /// 获取区域中的所有单词。
    /// </summary>
    IWordWords? Words { get; }

    /// <summary>
    /// 获取区域中的所有字符。
    /// </summary>
    IWordCharacters? Characters { get; }

    /// <summary>
    /// 获取区域中的所有表格。
    /// </summary>
    IWordTables? Tables { get; }

    /// <summary>
    /// 获取区域中的所有书签。
    /// </summary>
    IWordBookmarks? Bookmarks { get; }

    /// <summary>
    /// 获取区域中的所有字段。
    /// </summary>
    IWordFields? Fields { get; }

    /// <summary>
    /// 获取区域中的所有超链接。
    /// </summary>
    IWordHyperlinks? Hyperlinks { get; }

    /// <summary>
    /// 获取区域中的所有窗体字段。
    /// </summary>
    IWordFormFields? FormFields { get; }

    /// <summary>
    /// 获取区域中的所有修订。
    /// </summary>
    IWordRevisions? Revisions { get; }

    /// <summary>
    /// 获取区域中的所有注释。
    /// </summary>
    IWordComments? Comments { get; }

    /// <summary>
    /// 获取区域中的所有脚注。
    /// </summary>
    IWordFootnotes? Footnotes { get; }

    /// <summary>
    /// 获取区域中的所有尾注。
    /// </summary>
    IWordEndnotes? Endnotes { get; }


    #endregion

    #region 状态和工具属性 (State & Utility Properties)

    /// <summary>
    /// 获取或设置区域是否已进行拼写检查。
    /// </summary>
    bool SpellingChecked { get; set; }

    /// <summary>
    /// 获取或设置区域是否已进行语法检查。
    /// </summary>
    bool GrammarChecked { get; set; }

    /// <summary>
    /// 获取或设置拼写和语法检查器是否忽略指定的文本。
    /// </summary>
    int NoProofing { get; set; }

    /// <summary>
    /// 获取 Find 对象，用于查找操作。
    /// </summary>
    IWordFind? Find { get; }

    /// <summary>
    /// 获取或设置 FormattedText 对象，包含带格式的文本。
    /// </summary>
    IWordRange? FormattedText { get; set; }

    #endregion

    #region 基本方法 (Basic Methods)

    /// <summary>
    /// 选中此区域。
    /// </summary>
    void Select();

    /// <summary>
    /// 设置区域的起始和结束位置。
    /// </summary>
    /// <param name="start">起始位置。</param>
    /// <param name="end">结束位置。</param>
    void SetRange(int start, int end);

    /// <summary>
    /// 在当前范围之后插入指定文本。
    /// </summary>
    /// <param name="text">要插入的文本内容。</param>
    void InsertAfter(string text);

    /// <summary>
    /// 在当前范围之前插入指定文本。
    /// </summary>
    /// <param name="text">要插入的文本内容。</param>
    void InsertBefore(string text);

    /// <summary>
    /// 将范围折叠到指定方向。
    /// </summary>
    /// <param name="Direction">折叠方向，wdCollapseStart表示折叠到范围开始位置，wdCollapseEnd表示折叠到范围结束位置。</param>
    void Collapse(WdCollapseDirection Direction);

    /// <summary>
    /// 将范围折叠到默认方向（开始位置）。
    /// </summary>
    void Collapse();

    /// <summary>
    /// 在当前区域位置插入分隔符。
    /// <para>如果未指定分隔符类型，则默认插入分页符。</para>
    /// </summary>
    /// <param name="type">要插入的分隔符类型，参考 <see cref="WdBreakType"/> 枚举。</param>
    void InsertBreak(WdBreakType? type = null);

    /// <summary>
    /// 将当前范围的内容剪切到剪贴板。
    /// </summary>
    void Cut();

    /// <summary>
    /// 将指定范围的内容复制到剪贴板。
    /// </summary>
    void Copy();

    /// <summary>
    /// 将剪贴板的内容粘贴到区域中。
    /// </summary>
    void Paste();

    /// <summary>
    /// 删除区域中的内容。
    /// </summary>
    void Delete();


    #endregion

    #region 更多集合属性 (More Collection Properties)

    /// <summary>
    /// 获取区域中的所有形状范围。
    /// </summary>
    IWordShapeRange? ShapeRange { get; }

    /// <summary>
    /// 获取区域中的所有内嵌形状。
    /// </summary>
    IWordInlineShapes? InlineShapes { get; }

    /// <summary>
    /// 获取区域中的所有边框。
    /// </summary>
    IWordBorders? Borders { get; }

    /// <summary>
    /// 获取区域中的所有列表段落。
    /// </summary>
    IWordListParagraphs? ListParagraphs { get; }

    /// <summary>
    /// 获取区域中的可读性统计信息。
    /// </summary>
    IWordReadabilityStatistics? ReadabilityStatistics { get; }

    /// <summary>
    /// 获取区域中的拼写错误。
    /// </summary>
    IWordProofreadingErrors? SpellingErrors { get; }

    /// <summary>
    /// 获取区域中的语法错误。
    /// </summary>
    IWordProofreadingErrors? GrammaticalErrors { get; }

    /// <summary>
    /// 获取区域中的所有子文档。
    /// </summary>
    IWordSubdocuments? Subdocuments { get; }

    /// <summary>
    /// 获取区域中的内容控件。
    /// </summary>
    IWordContentControls? ContentControls { get; }

    /// <summary>
    /// 获取区域中的冲突。
    /// </summary>
    IWordConflicts? Conflicts { get; }

    /// <summary>
    /// 获取区域中的编辑者。
    /// </summary>
    IWordEditors? Editors { get; }

    #endregion

    #region 更多格式化属性 (More Formatting Properties)

    /// <summary>
    /// 获取或设置区域的双向粗体格式 (用于东亚语言)。
    /// </summary>
    int BoldBi { get; set; }

    /// <summary>
    /// 获取或设置区域的双向斜体格式 (用于东亚语言)。
    /// </summary>
    int ItalicBi { get; set; }

    /// <summary>
    /// 获取或设置区域的着重号。
    /// </summary>
    WdEmphasisMark EmphasisMark { get; set; }

    /// <summary>
    /// 获取或设置区域的字符宽度 (用于东亚语言)。
    /// </summary>
    WdCharacterWidth CharacterWidth { get; set; }

    /// <summary>
    /// 获取或设置区域的水平垂直文本格式 (用于东亚语言)。
    /// </summary>
    WdHorizontalInVerticalType HorizontalInVertical { get; set; }

    /// <summary>
    /// 获取或设置区域的文字方向。
    /// </summary>
    WdTextOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置区域的两行合一格式。
    /// </summary>
    WdTwoLinesInOneType TwoLinesInOne { get; set; }

    /// <summary>
    /// 获取或设置区域的语言ID。
    /// </summary>
    WdLanguageID LanguageID { get; set; }

    /// <summary>
    /// 获取或设置区域的东亚语言ID。
    /// </summary>
    WdLanguageID LanguageIDFarEast { get; set; }

    /// <summary>
    /// 获取或设置区域的其他语言ID。
    /// </summary>
    WdLanguageID LanguageIDOther { get; set; }

    /// <summary>
    /// 获取或设置是否已检测到区域的语言。
    /// </summary>
    bool LanguageDetected { get; set; }

    /// <summary>
    /// 获取或设置是否禁用字符网格。
    /// </summary>
    bool DisableCharacterSpaceGrid { get; set; }

    /// <summary>
    /// 获取或设置区域的ID (用于Web页)。
    /// </summary>
    string ID { get; set; }

    #endregion

    #region 更多方法 (More Methods)

    /// <summary>
    /// 执行拼写检查。
    /// </summary>
    void CheckSpelling();

    /// <summary>
    /// 执行语法检查。
    /// </summary>
    void CheckGrammar();

    #endregion

    /// <summary>
    /// 将范围折叠到指定方向（开始或结束位置）。
    /// </summary>
    /// <param name="direction">折叠方向，默认为wdCollapseStart（折叠到开始位置）。</param>
    void Collapse(WdCollapseDirection? direction = null);

    /// <summary>
    /// 获取指定单位的下一个区域范围。
    /// </summary>
    /// <param name="unit">移动单位，默认为wdWord（单词）。</param>
    /// <param name="count">移动的单位数量，默认为1。</param>
    /// <returns>返回一个新的范围对象，表示移动后的位置；如果操作失败，则返回null。</returns>
    IWordRange? Next(WdUnits? unit = null, int? count = null);

    /// <summary>
    /// 获取指定单位的上一个区域范围。
    /// </summary>
    /// <param name="unit">移动单位，默认为wdWord（单词）。</param>
    /// <param name="count">移动的单位数量，默认为1。</param>
    /// <returns>返回一个新的范围对象，表示移动后的位置；如果操作失败，则返回null。</returns>
    IWordRange? Previous(WdUnits? unit = null, int? count = null);

    /// <summary>
    /// 将范围的起始位置移动到指定单元的开始位置。
    /// </summary>
    /// <param name="unit">移动单位，默认为wdWord（单词）。</param>
    /// <param name="extend">指定移动方式，wdMove表示移动范围边界，wdExtend表示扩展范围；默认为wdMove。</param>
    /// <returns>返回移动的字符数；如果操作失败，则返回null。</returns>
    int? StartOf(WdUnits? unit = null, WdMovementType? extend = null);

    /// <summary>
    /// 将范围的结束位置移动到指定单元的结束位置。
    /// </summary>
    /// <param name="unit">移动单位，默认为wdWord（单词）。</param>
    /// <param name="extend">指定移动方式，wdMove表示移动范围边界，wdExtend表示扩展范围；默认为wdMove。</param>
    /// <returns>返回移动的字符数；如果操作失败，则返回null。</returns>
    int? EndOf(WdUnits? unit = null, WdMovementType? extend = null);

    /// <summary>
    /// 将范围移动指定的单位和数量。
    /// </summary>
    /// <param name="unit">移动单位，默认为wdCharacter（字符）。</param>
    /// <param name="count">移动的单位数量，默认为1。</param>
    /// <returns>返回实际移动的字符数；如果操作失败，则返回null。</returns>
    int? Move(WdUnits? unit = null, int? count = null);

    /// <summary>
    /// 移动范围的起始边界指定的单位和数量。
    /// </summary>
    /// <param name="unit">移动单位，默认为wdCharacter（字符）。</param>
    /// <param name="count">移动的单位数量，默认为1。</param>
    /// <returns>返回实际移动的字符数；如果操作失败，则返回null。</returns>
    int? MoveStart(WdUnits? unit = null, int? count = null);

    /// <summary>
    /// 移动范围的结束边界指定的单位和数量。
    /// </summary>
    /// <param name="unit">移动单位，默认为wdCharacter（字符）。</param>
    /// <param name="count">移动的单位数量，默认为1。</param>
    /// <returns>返回实际移动的字符数；如果操作失败，则返回null。</returns>
    int? MoveEnd(WdUnits? unit = null, int? count = null);

    /// <summary>
    /// 将范围移动，直到遇到不在指定字符集中的字符为止，或移动指定次数。
    /// </summary>
    /// <param name="cset">要移动的字符集（字符串）。</param>
    /// <param name="count">移动的单位数量，默认为wdForward（向前移动）。</param>
    /// <returns>返回实际移动的字符数；如果操作失败，则返回null。</returns>
    int? MoveWhile(string cset, int? count = null);

    /// <summary>
    /// 移动范围的起始边界，直到遇到不在指定字符集中的字符为止，或移动指定次数。
    /// </summary>
    /// <param name="cset">要移动的字符集（字符串）。</param>
    /// <param name="count">移动的单位数量，默认为wdForward（向前移动）。</param>
    /// <returns>返回实际移动的字符数；如果操作失败，则返回null。</returns>
    int? MoveStartWhile(string cset, int? count = null);

    /// <summary>
    /// 移动范围的结束边界，直到遇到不在指定字符集中的字符为止，或移动指定次数。
    /// </summary>
    /// <param name="cset">要移动的字符集（字符串）。</param>
    /// <param name="count">移动的单位数量，默认为wdForward（向前移动）。</param>
    /// <returns>返回实际移动的字符数；如果操作失败，则返回null。</returns>
    int? MoveEndWhile(string cset, int? count = null);

    /// <summary>
    /// 将范围移动，直到遇到指定字符集中的字符为止，或移动指定次数。
    /// </summary>
    /// <param name="cset">要查找的字符集（字符串）。</param>
    /// <param name="count">移动的单位数量，默认为wdForward（向前移动）。</param>
    /// <returns>返回实际移动的字符数；如果操作失败，则返回null。</returns>
    int? MoveUntil(string cset, int? count = null);

    /// <summary>
    /// 移动范围的起始边界，直到遇到指定字符集中的字符为止，或移动指定次数。
    /// </summary>
    /// <param name="cset">要查找的字符集（字符串）。</param>
    /// <param name="count">移动的单位数量，默认为wdForward（向前移动）。</param>
    /// <returns>返回实际移动的字符数；如果操作失败，则返回null。</returns>
    int? MoveStartUntil(string cset, int? count = null);

    /// <summary>
    /// 移动范围的结束边界，直到遇到指定字符集中的字符为止，或移动指定次数。
    /// </summary>
    /// <param name="cset">要查找的字符集（字符串）。</param>
    /// <param name="count">移动的单位数量，默认为wdForward（向前移动）。</param>
    /// <returns>返回实际移动的字符数；如果操作失败，则返回null。</returns>
    int? MoveEndUntil(string cset, int? count = null);

    /// <summary>
    /// 在当前范围中插入一个外部文件的内容。
    /// </summary>
    /// <param name="fileName">要插入的文件的完整路径名。</param>
    /// <param name="range">指定插入文件内容的范围，如果为null则在当前位置插入。</param>
    /// <param name="confirmConversions">是否在插入不同格式的文件时显示确认对话框。如果为null则使用Word默认设置。</param>
    /// <param name="link">是否将插入的文件作为链接插入。如果为null则使用Word默认设置。</param>
    /// <param name="attachment">是否将插入的文件作为附件插入。如果为null则使用Word默认设置。</param>
    void InsertFile(string fileName, object? range = null, bool? confirmConversions = null, bool? link = null, bool? attachment = null);

    /// <summary>
    /// 检查指定范围是否在当前范围的同一篇文章中。
    /// </summary>
    /// <param name="Range">要检查的范围。</param>
    /// <returns>如果指定范围在当前范围的同一篇文章中则返回true，否则返回false或null。</returns>
    bool? InStory(IWordRange Range);

    /// <summary>
    /// 检查指定范围是否在当前范围内。
    /// </summary>
    /// <param name="Range">要检查的范围。</param>
    /// <returns>如果指定范围在当前范围内则返回true，否则返回false或null。</returns>
    bool? InRange(IWordRange Range);

    /// <summary>
    /// 删除当前范围或指定单位的内容。
    /// </summary>
    /// <param name="unit">要删除的单位（如字符、单词、段落等），如果为null则删除当前范围的内容。</param>
    /// <param name="count">要删除的单位数量，默认为1。</param>
    /// <returns>返回实际删除的字符数，如果操作失败则返回null。</returns>
    int? Delete(WdUnits? unit = null, int? count = null);

    /// <summary>
    /// 选择整个文档的内容。
    /// </summary>
    void WholeStory();

    /// <summary>
    /// 扩展当前范围以包含指定单位的整个内容。
    /// </summary>
    /// <param name="unit">要扩展到的单位（如字符、单词、段落等），如果为null则使用默认单位。</param>
    /// <returns>返回扩展的字符数，如果操作失败则返回null。</returns>
    int? Expand(WdUnits? unit = null);

    /// <summary>
    /// 在当前范围的开始处插入一个段落标记。
    /// </summary>
    void InsertParagraph();

    /// <summary>
    /// 在当前范围的末尾插入一个段落标记。
    /// </summary>
    void InsertParagraphAfter();

    /// <summary>
    /// 在当前范围插入一个特殊符号。
    /// </summary>
    /// <param name="characterNumber">要插入的符号的字符代码。</param>
    /// <param name="font">符号的字体名称，如果为null则使用当前字体。</param>
    /// <param name="Unicode">是否使用Unicode编码，如果为null则使用默认设置。</param>
    /// <param name="Bias">字体选择的倾向性，如果为null则使用默认设置。</param>
    void InsertSymbol(int characterNumber, string? font = null, bool? Unicode = null, WdFontBias? Bias = null);

    /// <summary>
    /// 将当前范围的内容复制为图片格式到剪贴板。
    /// </summary>
    void CopyAsPicture();

    /// <summary>
    /// 对当前范围的内容进行升序排序。
    /// </summary>
    void SortAscending();

    /// <summary>
    /// 对当前范围的内容进行降序排序。
    /// </summary>
    void SortDescending();

    /// <summary>
    /// 检查指定范围是否与当前范围相等。
    /// </summary>
    /// <param name="range">要比较的范围。</param>
    /// <returns>如果两个范围相等则返回true，否则返回false或null。</returns>
    bool? IsEqual(IWordRange range);

    /// <summary>
    /// 计算当前范围中文本表示的数学表达式。
    /// </summary>
    /// <returns>返回计算结果，如果无法计算则返回null。</returns>
    float? Calculate();

    /// <summary>
    /// 跳转到文档中的指定位置
    /// </summary>
    /// <param name="what">要跳转到的对象类型，如页面、书签、行等</param>
    /// <param name="which">跳转方向，如下一个、上一个、绝对位置等</param>
    /// <param name="count">跳转的次数或位置</param>
    /// <param name="name">特定书签或项目的名称</param>
    /// <returns>表示跳转后位置的Word范围对象，如果未找到则返回null</returns>
    IWordRange? GoTo(WdGoToItem? what = null, WdGoToDirection? which = null, int? count = null, string? name = null);

    /// <summary>
    /// 跳转到下一个指定类型的项目
    /// </summary>
    /// <param name="what">要查找的项目类型</param>
    /// <returns>表示下一个项目位置的Word范围对象，如果未找到则返回null</returns>
    IWordRange? GoToNext(WdGoToItem what);

    /// <summary>
    /// 跳转到上一个指定类型的项目
    /// </summary>
    /// <param name="what">要查找的项目类型</param>
    /// <returns>表示上一个项目位置的Word范围对象，如果未找到则返回null</returns>
    IWordRange? GoToPrevious(WdGoToItem what);

    /// <summary>
    /// 以特殊格式粘贴剪贴板内容
    /// </summary>
    /// <param name="iconIndex">图标文件中的图标索引</param>
    /// <param name="link">是否链接到源文件</param>
    /// <param name="placement">OLE对象的放置方式</param>
    /// <param name="displayAsIcon">是否以图标形式显示</param>
    /// <param name="dataType">要粘贴的数据类型</param>
    /// <param name="iconFileName">图标的文件路径</param>
    /// <param name="IconLabel">图标的标签文本</param>
    void PasteSpecial(int? iconIndex = null, object? link = null, WdOLEPlacement? placement = null,
                    bool? displayAsIcon = null, WdPasteDataType? dataType = null,
                    string? iconFileName = null, string? IconLabel = null);


    /// <summary>
    /// 查找并显示名称属性信息
    /// </summary>
    void LookupNameProperties();

    /// <summary>
    /// 计算文档的统计信息
    /// </summary>
    /// <param name="Statistic">要计算的统计类型</param>
    /// <returns>计算得到的统计值，如果无法计算则返回null</returns>
    int? ComputeStatistics(WdStatistic Statistic);

    /// <summary>
    /// 重新定位当前范围
    /// </summary>
    /// <param name="direction">重新定位的方向</param>
    void Relocate([ConvertInt] WdRelocate direction);

    /// <summary>
    /// 检查所选文本的同义词
    /// </summary>
    void CheckSynonyms();

    /// <summary>
    /// 插入自动图文集条目
    /// </summary>
    void InsertAutoText();

    /// <summary>
    /// 自动格式化当前范围的内容
    /// </summary>
    void AutoFormat();

    /// <summary>
    /// 在当前位置前插入一个段落标记
    /// </summary>
    void InsertParagraphBefore();

    /// <summary>
    /// 移动到下一个子文档
    /// </summary>
    void NextSubdocument();

    /// <summary>
    /// 移动到上一个子文档
    /// </summary>
    void PreviousSubdocument();

    /// <summary>
    /// 将剪贴板内容作为嵌套表格粘贴
    /// </summary>
    void PasteAsNestedTable();


    /// <summary>
    /// 检测当前范围的语言
    /// </summary>
    void DetectLanguage();

    /// <summary>
    /// 执行拼写检查。
    /// </summary>
    /// <param name="customDictionary">自定义词典，用于扩展默认词典进行拼写检查</param>
    /// <param name="ignoreUppercase">是否忽略大写单词</param>
    /// <param name="alwaysSuggest">是否始终提供建议</param>
    /// <param name="customDictionary2">附加自定义词典2</param>
    /// <param name="customDictionary3">附加自定义词典3</param>
    /// <param name="customDictionary4">附加自定义词典4</param>
    /// <param name="customDictionary5">附加自定义词典5</param>
    /// <param name="customDictionary6">附加自定义词典6</param>
    /// <param name="customDictionary7">附加自定义词典7</param>
    /// <param name="customDictionary8">附加自定义词典8</param>
    /// <param name="customDictionary9">附加自定义词典9</param>
    /// <param name="customDictionary10">附加自定义词典10</param>
    void CheckSpelling(IWordDictionary? customDictionary = null, bool? ignoreUppercase = null, bool? alwaysSuggest = null,
                        IWordDictionary? customDictionary2 = null, IWordDictionary? customDictionary3 = null,
                        IWordDictionary? customDictionary4 = null, IWordDictionary? customDictionary5 = null,
                        IWordDictionary? customDictionary6 = null, IWordDictionary? customDictionary7 = null,
                        IWordDictionary? customDictionary8 = null, IWordDictionary? customDictionary9 = null,
                        IWordDictionary? customDictionary10 = null);

    /// <summary>
    /// 获取拼写建议。
    /// </summary>
    /// <param name="CustomDictionary">自定义词典，用于扩展默认词典</param>
    /// <param name="ignoreUppercase">是否忽略大写单词</param>
    /// <param name="MainDictionary">主词典</param>
    /// <param name="SuggestionMode">建议模式，指定如何生成建议</param>
    /// <param name="customDictionary2">附加自定义词典2</param>
    /// <param name="customDictionary3">附加自定义词典3</param>
    /// <param name="customDictionary4">附加自定义词典4</param>
    /// <param name="customDictionary5">附加自定义词典5</param>
    /// <param name="customDictionary6">附加自定义词典6</param>
    /// <param name="customDictionary7">附加自定义词典7</param>
    /// <param name="customDictionary8">附加自定义词典8</param>
    /// <param name="customDictionary9">附加自定义词典9</param>
    /// <param name="customDictionary10">附加自定义词典10</param>
    /// <returns>返回拼写建议列表，如果没有找到建议则返回null</returns>
    IWordSpellingSuggestions? GetSpellingSuggestions(IWordDictionary? CustomDictionary = null, bool? ignoreUppercase = null, IWordDictionary? MainDictionary = null,
                        WdSpellingWordType? SuggestionMode = null, IWordDictionary? customDictionary2 = null, IWordDictionary? customDictionary3 = null,
                        IWordDictionary? customDictionary4 = null, IWordDictionary? customDictionary5 = null,
                        IWordDictionary? customDictionary6 = null, IWordDictionary? customDictionary7 = null,
                        IWordDictionary? customDictionary8 = null, IWordDictionary? customDictionary9 = null,
                        IWordDictionary? customDictionary10 = null);

    /// <summary>
    /// 插入数据库内容。
    /// </summary>
    /// <param name="format">表格格式</param>
    /// <param name="style">样式</param>
    /// <param name="linkToSource">是否链接到源</param>
    /// <param name="connection">数据库连接信息</param>
    /// <param name="SQLStatement">SQL语句</param>
    /// <param name="SQLStatement1">附加SQL语句</param>
    /// <param name="passwordDocument">文档密码</param>
    /// <param name="passwordTemplate">模板密码</param>
    /// <param name="writePasswordDocument">写入文档密码</param>
    /// <param name="writePasswordTemplate">写入模板密码</param>
    /// <param name="dataSource">数据源</param>
    /// <param name="from">起始记录位置</param>
    /// <param name="to">结束记录位置</param>
    /// <param name="includeFields">是否包含字段</param>
    void InsertDatabase(WdTableFormat? format = null, object? style = null, bool? linkToSource = null,
                        object? connection = null, string? SQLStatement = null, string? SQLStatement1 = null,
                        string? passwordDocument = null, string? passwordTemplate = null, string? writePasswordDocument = null,
                        string? writePasswordTemplate = null, string? dataSource = null, int? from = null, int? to = null, bool? includeFields = null);

    /// <summary>
    /// 转换韩文和汉字。
    /// </summary>
    /// <param name="conversionsMode">转换模式</param>
    /// <param name="fastConversion">是否快速转换</param>
    /// <param name="checkHangulEnding">是否检查韩文结尾</param>
    /// <param name="enableRecentOrdering">是否启用最近排序</param>
    /// <param name="customDictionary">自定义词典</param>
    void ConvertHangulAndHanja(WdMultipleWordConversionsMode conversionsMode, bool? fastConversion, bool? checkHangulEnding,
                               bool? enableRecentOrdering, IWordDictionary? customDictionary);

    /// <summary>
    /// 修改封装样式。
    /// </summary>
    /// <param name="style">封装样式</param>
    /// <param name="symbol">封装符号类型</param>
    /// <param name="enclosedText">被封装的文本，如果为null则使用当前范围的文本</param>
    void ModifyEnclosure(WdEncloseStyle? style, WdEnclosureType symbol, string? enclosedText = null);

    /// <summary>
    /// 添加拼音指南。
    /// </summary>
    /// <param name="Text">要添加拼音的文本</param>
    /// <param name="alignment">对齐方式</param>
    /// <param name="raise">拼音提升高度</param>
    /// <param name="fontSize">拼音字体大小</param>
    /// <param name="fontName">拼音字体名称</param>
    void PhoneticGuide(string Text, WdPhoneticGuideAlignmentType alignment = WdPhoneticGuideAlignmentType.wdPhoneticGuideAlignmentLeft,
                       int raise = 0, int fontSize = 0, string fontName = "");

    /// <summary>
    /// 插入日期和时间。
    /// </summary>
    /// <param name="dateTimeFormat">日期时间格式</param>
    /// <param name="insertAsField">是否作为字段插入</param>
    /// <param name="insertAsFullWidth">是否以全角格式插入</param>
    /// <param name="dateLanguage">日期语言</param>
    /// <param name="calendarType">日历类型</param>
    void InsertDateTime(string? dateTimeFormat = null, bool? insertAsField = null, bool? insertAsFullWidth = null,
                        WdDateLanguage? dateLanguage = null, WdCalendarTypeBi? calendarType = null);

    /// <summary>
    /// 对范围内的内容进行排序。
    /// </summary>
    /// <param name="excludeHeader">是否排除标题行</param>
    /// <param name="fieldNumber">第一个排序字段</param>
    /// <param name="sortFieldType">第一个排序字段类型</param>
    /// <param name="sortOrder">第一个排序顺序</param>
    /// <param name="fieldNumber2">第二个排序字段</param>
    /// <param name="sortFieldType2">第二个排序字段类型</param>
    /// <param name="sortOrder2">第二个排序顺序</param>
    /// <param name="fieldNumber3">第三个排序字段</param>
    /// <param name="sortFieldType3">第三个排序字段类型</param>
    /// <param name="sortOrder3">第三个排序顺序</param>
    /// <param name="sortColumn">排序列</param>
    /// <param name="separator">分隔符类型</param>
    /// <param name="caseSensitive">是否区分大小写</param>
    /// <param name="bidiSort">是否双向排序</param>
    /// <param name="ignoreThe">是否忽略英文"the"</param>
    /// <param name="ignoreKashida">是否忽略阿拉伯语kashida</param>
    /// <param name="ignoreDiacritics">是否忽略变音符号</param>
    /// <param name="ignoreHe">是否忽略希伯来语he</param>
    /// <param name="languageID">语言ID</param>
    void Sort(bool? excludeHeader = null, string? fieldNumber = null, WdSortFieldType? sortFieldType = null,
            WdSortOrder? sortOrder = null, string? fieldNumber2 = null, WdSortFieldType? sortFieldType2 = null,
            WdSortOrder? sortOrder2 = null, string? fieldNumber3 = null, WdSortFieldType? sortFieldType3 = null,
            WdSortOrder? sortOrder3 = null, IWordRange? sortColumn = null, WdSortSeparator? separator = null,
            bool? caseSensitive = null, bool? bidiSort = null, bool? ignoreThe = null,
            bool? ignoreKashida = null, bool? ignoreDiacritics = null, bool? ignoreHe = null,
            WdLanguageID? languageID = null);

    /// <summary>
    /// 对范围内的内容按标题进行排序。
    /// </summary>
    /// <param name="sortFieldType">排序字段类型</param>
    /// <param name="sortOrder">排序顺序</param>
    /// <param name="caseSensitive">是否区分大小写</param>
    /// <param name="bidiSort">是否双向排序</param>
    /// <param name="ignoreThe">是否忽略英文"the"</param>
    /// <param name="ignoreKashida">是否忽略阿拉伯语kashida</param>
    /// <param name="ignoreDiacritics">是否忽略变音符号</param>
    /// <param name="ignoreHe">是否忽略希伯来语he</param>
    /// <param name="languageID">语言ID</param>
    void SortByHeadings(WdSortFieldType? sortFieldType, WdSortOrder? sortOrder,
                        bool? caseSensitive, bool? bidiSort,
                        bool? ignoreThe, bool? ignoreKashida, bool? ignoreDiacritics,
                        bool? ignoreHe, WdLanguageID? languageID);


    /// <summary>
    /// 将当前范围转换为表格
    /// </summary>
    /// <param name="separator">表格字段分隔符，默认为null</param>
    /// <param name="numRows">表格行数，默认为null</param>
    /// <param name="numColumns">表格列数，默认为null</param>
    /// <param name="initialColumnWidth">初始列宽，默认为null</param>
    /// <param name="format">表格格式，默认为null</param>
    /// <param name="applyBorders">是否应用边框，默认为null</param>
    /// <param name="applyShading">是否应用底纹，默认为null</param>
    /// <param name="applyFont">是否应用字体格式，默认为null</param>
    /// <param name="applyColor">是否应用颜色，默认为null</param>
    /// <param name="applyHeadingRows">是否应用标题行格式，默认为null</param>
    /// <param name="applyLastRow">是否应用最后一行格式，默认为null</param>
    /// <param name="applyFirstColumn">是否应用第一列格式，默认为null</param>
    /// <param name="applyLastColumn">是否应用最后一列格式，默认为null</param>
    /// <param name="autoFit">是否自动调整大小，默认为null</param>
    /// <param name="autoFitBehavior">自动调整行为，默认为null</param>
    /// <param name="defaultTableBehavior">默认表格行为，默认为null</param>
    /// <returns>返回表示表格的IWordTable对象，如果转换失败则返回null</returns>
    IWordTable? ConvertToTable(WdTableFieldSeparator? separator = null, int? numRows = null, int? numColumns = null,
                                double? initialColumnWidth = null, WdTableFormat? format = null, bool? applyBorders = null,
                                bool? applyShading = null, bool? applyFont = null, bool? applyColor = null,
                                bool? applyHeadingRows = null, bool? applyLastRow = null, bool? applyFirstColumn = null,
                                bool? applyLastColumn = null, bool? autoFit = null, WdAutoFitBehavior? autoFitBehavior = null,
                                WdDefaultTableBehavior? defaultTableBehavior = null);

    /// <summary>
    /// 执行TCSC（繁简转换）转换
    /// </summary>
    /// <param name="WdTCSCConverterDirection">转换方向，默认为自动</param>
    /// <param name="commonTerms">是否转换常用词，默认为false</param>
    /// <param name="useVariants">是否使用变体，默认为false</param>
    void TCSCConverter(WdTCSCConverterDirection WdTCSCConverterDirection = WdTCSCConverterDirection.wdTCSCConverterDirectionAuto, bool commonTerms = false, bool useVariants = false);


    /// <summary>
    /// 粘贴并格式化内容
    /// </summary>
    /// <param name="type">恢复类型</param>
    void PasteAndFormat(WdRecoveryType type);

    /// <summary>
    /// 粘贴Excel表格
    /// </summary>
    /// <param name="linkedToExcel">是否链接到Excel</param>
    /// <param name="wordFormatting">是否使用Word格式</param>
    /// <param name="RTF">是否使用RTF格式</param>
    void PasteExcelTable(bool linkedToExcel, bool wordFormatting, bool RTF);

    /// <summary>
    /// 粘贴并追加表格
    /// </summary>
    void PasteAppendTable();

    /// <summary>
    /// 跳转到可编辑范围
    /// </summary>
    /// <param name="editorID">编辑器ID</param>
    /// <returns>返回表示可编辑范围的IWordRange对象，如果找不到则返回null</returns>
    IWordRange? GoToEditableRange(string? editorID);

    /// <summary>
    /// 插入XML内容
    /// </summary>
    /// <param name="XML">要插入的XML字符串</param>
    /// <param name="transform">转换样式，可选</param>
    void InsertXML(string XML, string? transform);

    /// <summary>
    /// 插入题注
    /// </summary>
    /// <param name="label">题注标签ID，默认为null</param>
    /// <param name="title">题注标题，默认为null</param>
    /// <param name="titleAutoText">自动图文集中的标题，默认为null</param>
    /// <param name="position">题注位置，默认为null</param>
    /// <param name="excludeLabel">是否排除标签，默认为null</param>
    void InsertCaption(WdCaptionLabelID? label = null, string? title = null, string? titleAutoText = null, WdCaptionPosition? position = null, bool? excludeLabel = null);

    /// <summary>
    /// 插入交叉引用（使用题注标签ID）
    /// </summary>
    /// <param name="referenceType">引用类型（题注标签ID）</param>
    /// <param name="referenceKind">引用种类</param>
    /// <param name="referenceItem">引用项目，默认为null</param>
    /// <param name="insertAsHyperlink">是否作为超链接插入，默认为null</param>
    /// <param name="includePosition">是否包含位置信息，默认为null</param>
    /// <param name="separateNumbers">是否分离数字，默认为null</param>
    /// <param name="separatorString">分隔字符串，默认为null</param>
    void InsertCrossReference(WdCaptionLabelID referenceType, WdReferenceKind referenceKind,
                              string? referenceItem = null, bool? insertAsHyperlink = null, bool? includePosition = null,
                              bool? separateNumbers = null, string? separatorString = null);

    /// <summary>
    /// 插入交叉引用（使用引用类型）
    /// </summary>
    ///<param name="referenceType">引用类型</param>
    /// <param name="referenceKind">引用种类</param>
    /// <param name="referenceItem">引用项目，默认为null</param>
    /// <param name="insertAsHyperlink">是否作为超链接插入，默认为null</param>
    /// <param name="includePosition">是否包含位置信息，默认为null</param>
    /// <param name="separateNumbers">是否分离数字，默认为null</param>
    /// <param name="separatorString">分隔字符串，默认为null</param>
    void InsertCrossReference(WdReferenceType referenceType, WdReferenceKind referenceKind,
                                string? referenceItem = null, bool? insertAsHyperlink = null, bool? includePosition = null,
                                bool? separateNumbers = null, string? separatorString = null);

    /// <summary>
    /// 导出文档片段到指定格式的文件
    /// </summary>
    /// <param name="fileName">输出文件名</param>
    /// <param name="format">保存格式</param>
    void ExportFragment(string fileName, WdSaveFormat format);

    /// <summary>
    /// 设置列表级别
    /// </summary>
    /// <param name="level">列表级别</param>
    void SetListLevel(short level);

    /// <summary>
    /// 插入对齐制表符
    /// </summary>
    /// <param name="alignment">对齐方式</param>
    /// <param name="relativeTo">相对于的位置，默认为0</param>
    void InsertAlignmentTab([ConvertInt] WdAlignmentTabAlignment alignment, int relativeTo = 0);

    /// <summary>
    /// 导入文档片段
    /// </summary>
    /// <param name="fileName">要导入的文件名</param>
    /// <param name="matchDestination">是否匹配目标格式，默认为false</param>
    void ImportFragment(string fileName, bool matchDestination = false);

    /// <summary>
    /// 导出为固定格式文件（如PDF或XPS）
    /// </summary>
    /// <param name="outputFileName">输出文件名</param>
    /// <param name="exportFormat">导出格式</param>
    /// <param name="openAfterExport">导出后是否打开文件，默认为false</param>
    /// <param name="optimizeFor">优化目标，默认为打印优化</param>
    /// <param name="exportCurrentPage">是否仅导出当前页，默认为false</param>
    /// <param name="item">要导出的项目类型，默认为文档内容</param>
    /// <param name="includeDocProps">是否包含文档属性，默认为false</param>
    /// <param name="keepIRM">是否保留IRM信息，默认为true</param>
    /// <param name="createBookmarks">创建书签的方式，默认为不创建</param>
    /// <param name="docStructureTags">是否包含文档结构标签，默认为true</param>
    /// <param name="bitmapMissingFonts">是否位图化缺失字体，默认为true</param>
    /// <param name="useISO19005_1">是否使用ISO 19005-1标准（PDF/A），默认为false</param>
    /// <param name="fixedFormatExtClassPtr">固定格式扩展类指针，默认为null</param>
    void ExportAsFixedFormat(string outputFileName, WdExportFormat exportFormat, bool openAfterExport = false,
                               WdExportOptimizeFor optimizeFor = WdExportOptimizeFor.wdExportOptimizeForPrint,
                               bool exportCurrentPage = false, WdExportItem item = WdExportItem.wdExportDocumentContent,
                               bool includeDocProps = false, bool keepIRM = true,
                               WdExportCreateBookmarks createBookmarks = WdExportCreateBookmarks.wdExportCreateNoBookmarks,
                               bool docStructureTags = true, bool bitmapMissingFonts = true,
                               bool useISO19005_1 = false, object? fixedFormatExtClassPtr = null);

    /// <summary>
    /// 获取或设置适合文本的宽度（以磅为单位）
    /// </summary>
    float FitTextWidth { get; set; }

    /// <summary>
    /// 获取或设置是否合并字符
    /// </summary>
    bool CombineCharacters { get; set; }

    /// <summary>
    /// 获取或设置是否显示所有内容（包括隐藏内容）
    /// </summary>
    bool ShowAll { get; set; }

    /// <summary>
    /// 获取顶级表格集合
    /// </summary>
    IWordTables? TopLevelTables { get; }

    /// <summary>
    /// 获取HTML分隔集合
    /// </summary>
    IWordHTMLDivisions? HTMLDivisions { get; }

    /// <summary>
    /// 获取Office脚本集合
    /// </summary>
    IOfficeScripts? Scripts { get; }

    /// <summary>
    /// 获取智能标签集合
    /// </summary>
    IWordSmartTags? SmartTags { get; }

    /// <summary>
    /// 获取脚注选项
    /// </summary>
    IWordFootnoteOptions? FootnoteOptions { get; }

    /// <summary>
    /// 获取尾注选项
    /// </summary>
    IWordEndnoteOptions? EndnoteOptions { get; }

    /// <summary>
    /// 获取XML节点集合
    /// </summary>
    IWordXMLNodes? XMLNodes { get; }

    /// <summary>
    /// 获取XML父节点
    /// </summary>
    IWordXMLNode? XMLParentNode { get; }

    /// <summary>
    /// 获取数学对象集合
    /// </summary>
    IWordOMaths? OMaths { get; }

    /// <summary>
    /// 获取协作锁定集合
    /// </summary>
    IWordCoAuthLocks? Locks { get; }

    /// <summary>
    /// 获取协作更新集合
    /// </summary>
    IWordCoAuthUpdates? Updates { get; }

    /// <summary>
    /// 获取屏幕上可见文本的数量
    /// </summary>
    int TextVisibleOnScreen { get; }

    /// <summary>
    /// 获取父内容控件
    /// </summary>
    IWordContentControl ParentContentControl { get; }

    /// <summary>
    /// 获取范围的XML表示
    /// </summary>
    string XML { get; }

    /// <summary>
    /// 获取范围的Word Open XML表示
    /// </summary>
    string WordOpenXML { get; }

    /// <summary>
    /// 获取增强图元文件位
    /// </summary>
    object EnhMetaFileBits { get; }

    /// <summary>
    /// 获取或设置字符样式
    /// </summary>
    object CharacterStyle { get; }

    /// <summary>
    /// 获取或设置段落样式
    /// </summary>
    object ParagraphStyle { get; }

    /// <summary>
    /// 获取或设置列表样式
    /// </summary>
    object ListStyle { get; }

    /// <summary>
    /// 获取或设置表格样式
    /// </summary>
    object TableStyle { get; }
}