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
public interface IWordRange : IDisposable
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
    /// 跳转到文档中的指定位置。
    /// </summary>
    /// <param name="what">要跳转到的项目类型，如书签、页面、行等，参考 <see cref="WdGoToItem"/> 枚举。</param>
    /// <param name="which">跳转方向或特定位置，参考 <see cref="WdGoToDirection"/> 枚举。</param>
    /// <param name="count">跳转的次数或位置编号，如果为null则使用默认值。</param>
    /// <param name="name">特定项目的名称（如书签名称），如果为null则使用默认值。</param>
    /// <returns>返回表示跳转位置的范围对象；如果操作失败，则返回null。</returns>
    IWordRange? GoTo(WdGoToItem? what = null, WdGoToDirection? which = null, int? count = null, string? name = null);

    /// <summary>
    /// 跳转到文档中下一个指定类型的项目。
    /// </summary>
    /// <param name="what">要查找的项目类型，参考 <see cref="WdGoToItem"/> 枚举。</param>
    /// <returns>返回表示下一个项目位置的范围对象；如果未找到，则返回null。</returns>
    IWordRange? GoToNext(WdGoToItem what);

    /// <summary>
    /// 跳转到文档中上一个指定类型的项目。
    /// </summary>
    /// <param name="what">要查找的项目类型，参考 <see cref="WdGoToItem"/> 枚举。</param>
    /// <returns>返回表示上一个项目位置的范围对象；如果未找到，则返回null。</returns>
    IWordRange? GoToPrevious(WdGoToItem what);

    /// <summary>
    /// 以特殊格式粘贴剪贴板内容。
    /// </summary>
    /// <param name="iconIndex">图标索引，如果为null则使用默认值。</param>
    /// <param name="link">是否链接到原始数据，如果为null则使用默认值。</param>
    /// <param name="placement">OLE对象的放置方式，参考 <see cref="WdOLEPlacement"/> 枚举。</param>
    /// <param name="displayAsIcon">是否以图标形式显示，如果为null则使用默认值。</param>
    /// <param name="dataType">粘贴数据的类型，参考 <see cref="WdPasteDataType"/> 枚举。</param>
    /// <param name="iconFileName">图标的文件名，如果为null则使用默认图标。</param>
    /// <param name="IconLabel">图标的标签文本，如果为null则使用默认标签。</param>
    void PasteSpecial(int? iconIndex = null, object? link = null, WdOLEPlacement? placement = null,
                    bool? displayAsIcon = null, WdPasteDataType? dataType = null,
                    string? iconFileName = null, string? IconLabel = null);

    /// <summary>
    /// 查找名称属性。
    /// </summary>
    void LookupNameProperties();

    /// <summary>
    /// 计算指定的文档统计信息。
    /// </summary>
    /// <param name="Statistic">要计算的统计类型，参考 <see cref="WdStatistic"/> 枚举。</param>
    /// <returns>返回计算结果的数值；如果操作失败，则返回null。</returns>
    int? ComputeStatistics(WdStatistic Statistic);

    /// <summary>
    /// 重新定位到指定方向。
    /// </summary>
    /// <param name="direction">重定位的方向，参考 <see cref="WdRelocate"/> 枚举。</param>
    void Relocate([ConvertInt] WdRelocate direction);

    /// <summary>
    /// 检查同义词。
    /// </summary>
    void CheckSynonyms();

    /// <summary>
    /// 插入自动图文集条目。
    /// </summary>
    void InsertAutoText();

    /// <summary>
    /// 自动格式化范围内容。
    /// </summary>
    void AutoFormat();

    /// <summary>
    /// 在当前范围前插入一个段落标记。
    /// </summary>
    void InsertParagraphBefore();

    /// <summary>
    /// 跳转到下一个子文档。
    /// </summary>
    void NextSubdocument();

    /// <summary>
    /// 跳转到上一个子文档。
    /// </summary>
    void PreviousSubdocument();

    /// <summary>
    /// 以嵌套表格形式粘贴内容。
    /// </summary>
    void PasteAsNestedTable();

    /// <summary>
    /// 检测范围内的语言。
    /// </summary>
    void DetectLanguage();

    /// <summary>
    /// 执行拼写检查。
    /// </summary>
    void CheckSpelling(IWordDictionary? customDictionary = null, bool? ignoreUppercase = null, bool? alwaysSuggest = null,
                        IWordDictionary? customDictionary2 = null, IWordDictionary? customDictionary3 = null,
                        IWordDictionary? customDictionary4 = null, IWordDictionary? customDictionary5 = null,
                        IWordDictionary? customDictionary6 = null, IWordDictionary? customDictionary7 = null,
                        IWordDictionary? customDictionary8 = null, IWordDictionary? customDictionary9 = null,
                        IWordDictionary? customDictionary10 = null);

    IWordSpellingSuggestions? GetSpellingSuggestions(IWordDictionary? CustomDictionary = null, bool? ignoreUppercase = null, IWordDictionary? MainDictionary = null,
                        WdSpellingWordType? SuggestionMode = null, IWordDictionary? customDictionary2 = null, IWordDictionary? customDictionary3 = null,
                        IWordDictionary? customDictionary4 = null, IWordDictionary? customDictionary5 = null,
                        IWordDictionary? customDictionary6 = null, IWordDictionary? customDictionary7 = null,
                        IWordDictionary? customDictionary8 = null, IWordDictionary? customDictionary9 = null,
                        IWordDictionary? customDictionary10 = null);



    void InsertDatabase(WdTableFormat? format = null, object? style = null, bool? linkToSource = null,
                        object? connection = null, string? SQLStatement = null, string? SQLStatement1 = null,
                        string? passwordDocument = null, string? passwordTemplate = null, string? writePasswordDocument = null,
                        string? writePasswordTemplate = null, string? dataSource = null, int? from = null, int? to = null, bool? includeFields = null);


    void ConvertHangulAndHanja(WdMultipleWordConversionsMode conversionsMode, bool? fastConversion, bool? checkHangulEnding,
                               bool? enableRecentOrdering, IWordDictionary? customDictionary);

    void ModifyEnclosure(WdEncloseStyle? style, WdEnclosureType symbol, string? enclosedText = null);

    void PhoneticGuide(string Text, WdPhoneticGuideAlignmentType alignment = WdPhoneticGuideAlignmentType.wdPhoneticGuideAlignmentLeft,
                       int raise = 0, int fontSize = 0, string fontName = "");


    void InsertDateTime(string? dateTimeFormat = null, bool? insertAsField = null, bool? insertAsFullWidth = null,
                        WdDateLanguage? dateLanguage = null, WdCalendarTypeBi? calendarType = null);


    void Sort(bool? excludeHeader = null, string? fieldNumber = null, WdSortFieldType? sortFieldType = null,
            WdSortOrder? sortOrder = null, string? fieldNumber2 = null, WdSortFieldType? sortFieldType2 = null,
            WdSortOrder? sortOrder2 = null, string? fieldNumber3 = null, WdSortFieldType? sortFieldType3 = null,
            WdSortOrder? sortOrder3 = null, IWordRange? sortColumn = null, WdSortSeparator? separator = null,
            bool? caseSensitive = null, bool? bidiSort = null, bool? ignoreThe = null,
            bool? ignoreKashida = null, bool? ignoreDiacritics = null, bool? ignoreHe = null,
            WdLanguageID? languageID = null);

    void SortByHeadings(WdSortFieldType? sortFieldType, WdSortOrder? sortOrder,
                        bool? caseSensitive, bool? bidiSort,
                        bool? ignoreThe, bool? ignoreKashida, bool? ignoreDiacritics,
                        bool? ignoreHe, WdLanguageID? languageID);


    IWordTable? ConvertToTable(WdTableFieldSeparator? separator = null, int? numRows = null, int? numColumns = null,
                                double? initialColumnWidth = null, WdTableFormat? format = null, bool? applyBorders = null,
                                bool? applyShading = null, bool? applyFont = null, bool? applyColor = null,
                                bool? applyHeadingRows = null, bool? applyLastRow = null, bool? applyFirstColumn = null,
                                bool? applyLastColumn = null, bool? autoFit = null, WdAutoFitBehavior? autoFitBehavior = null,
                                WdDefaultTableBehavior? defaultTableBehavior = null);

    void TCSCConverter(WdTCSCConverterDirection WdTCSCConverterDirection = WdTCSCConverterDirection.wdTCSCConverterDirectionAuto, bool commonTerms = false, bool useVariants = false);


    void PasteAndFormat(WdRecoveryType type);

    void PasteExcelTable(bool linkedToExcel, bool wordFormatting, bool RTF);

    void PasteAppendTable();

    IWordRange? GoToEditableRange(string? editorID);

    void InsertXML(string XML, string? transform);

    void InsertCaption(WdCaptionLabelID? label = null, string? title = null, string? titleAutoText = null, WdCaptionPosition? position = null, bool? excludeLabel = null);

    void InsertCrossReference(WdCaptionLabelID referenceType, WdReferenceKind referenceKind,
                              string? referenceItem = null, bool? insertAsHyperlink = null, bool? includePosition = null,
                              bool? separateNumbers = null, string? separatorString = null);

    void InsertCrossReference(WdReferenceType referenceType, WdReferenceKind referenceKind,
                                string? referenceItem = null, bool? insertAsHyperlink = null, bool? includePosition = null,
                                bool? separateNumbers = null, string? separatorString = null);

    void ExportFragment(string fileName, WdSaveFormat format);

    void SetListLevel(short level);

    void InsertAlignmentTab([ConvertInt] WdAlignmentTabAlignment alignment, int relativeTo = 0);

    void ImportFragment(string fileName, bool matchDestination = false);

    void ExportAsFixedFormat(string outputFileName, WdExportFormat exportFormat, bool openAfterExport = false,
                               WdExportOptimizeFor optimizeFor = WdExportOptimizeFor.wdExportOptimizeForPrint,
                               bool exportCurrentPage = false, WdExportItem item = WdExportItem.wdExportDocumentContent,
                               bool includeDocProps = false, bool keepIRM = true,
                               WdExportCreateBookmarks createBookmarks = WdExportCreateBookmarks.wdExportCreateNoBookmarks,
                               bool docStructureTags = true, bool bitmapMissingFonts = true,
                               bool useISO19005_1 = false, object? fixedFormatExtClassPtr = null);


    float FitTextWidth { get; set; }

    bool CombineCharacters { get; set; }

    bool ShowAll { get; set; }

    IWordTables? TopLevelTables { get; }

    IWordHTMLDivisions? HTMLDivisions { get; }

    IOfficeScripts? Scripts { get; }

    IWordSmartTags? SmartTags { get; }

    IWordFootnoteOptions? FootnoteOptions { get; }

    IWordEndnoteOptions? EndnoteOptions { get; }

    IWordXMLNodes? XMLNodes { get; }

    IWordXMLNode? XMLParentNode { get; }

    IWordOMaths? OMaths { get; }

    IWordCoAuthLocks? Locks { get; }

    IWordCoAuthUpdates? Updates { get; }

    int TextVisibleOnScreen { get; }

    IWordContentControl ParentContentControl { get; }

    string XML { get; }

    string WordOpenXML { get; }

    object EnhMetaFileBits { get; }

    object CharacterStyle { get; }

    object ParagraphStyle { get; }

    object ListStyle { get; }

    object TableStyle { get; }
}