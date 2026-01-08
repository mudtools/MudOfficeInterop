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
    /// 获取或设置指定范围内的文本。
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置一个Range对象，表示指定范围内格式化文本。
    /// </summary>
    IWordRange? FormattedText { get; set; }

    /// <summary>
    /// 获取或设置范围的起始字符位置。
    /// </summary>
    int Start { get; set; }

    /// <summary>
    /// 获取或设置范围的结束字符位置。
    /// </summary>
    int End { get; set; }

    /// <summary>
    /// 获取或设置表示指定对象字符格式设置的Font对象。
    /// </summary>
    IWordFont? Font { get; set; }

    /// <summary>
    /// 获取表示指定范围所有属性的Range对象。
    /// </summary>
    IWordRange? Duplicate { get; }

    /// <summary>
    /// 获取指定范围的文章类型。
    /// </summary>
    WdStoryType StoryType { get; }

    /// <summary>
    /// 获取表示指定范围内所有表格的Tables集合。
    /// </summary>
    IWordTables? Tables { get; }

    /// <summary>
    /// 获取表示范围内所有单词的Words集合。
    /// </summary>
    IWordWords? Words { get; }

    /// <summary>
    /// 获取表示范围内所有句子的Sentences集合。
    /// </summary>
    IWordSentences? Sentences { get; }

    /// <summary>
    /// 获取表示范围内所有字符的Characters集合。
    /// </summary>
    IWordCharacters? Characters { get; }

    /// <summary>
    /// 获取表示范围内所有脚注的Footnotes集合。
    /// </summary>
    IWordFootnotes? Footnotes { get; }

    /// <summary>
    /// 获取表示范围内所有尾注的Endnotes集合。
    /// </summary>
    IWordEndnotes? Endnotes { get; }

    /// <summary>
    /// 获取表示指定范围内所有批注的Comments集合。
    /// </summary>
    IWordComments? Comments { get; }

    /// <summary>
    /// 获取表示范围内所有表格单元格的Cells集合。
    /// </summary>
    IWordCells? Cells { get; }

    /// <summary>
    /// 获取表示指定范围内所有节的Sections集合。
    /// </summary>
    IWordSections? Sections { get; }

    /// <summary>
    /// 获取表示指定范围内所有段落的Paragraphs集合。
    /// </summary>
    IWordParagraphs? Paragraphs { get; }

    /// <summary>
    /// 获取或设置表示指定对象所有边框的Borders集合。
    /// </summary>
    IWordBorders? Borders { get; set; }

    /// <summary>
    /// 获取表示指定对象底纹格式设置的Shading对象。
    /// </summary>
    IWordShading? Shading { get; }

    /// <summary>
    /// 获取或设置控制如何从指定范围检索文本的TextRetrievalMode对象。
    /// </summary>
    IWordTextRetrievalMode? TextRetrievalMode { get; set; }

    /// <summary>
    /// 获取表示范围内所有字段的只读Fields集合。
    /// </summary>
    IWordFields? Fields { get; }

    /// <summary>
    /// 获取表示范围内所有表单字段的FormFields集合。
    /// </summary>
    IWordFormFields? FormFields { get; }

    /// <summary>
    /// 获取表示范围内所有框架的Frames集合。
    /// </summary>
    IWordFrames? Frames { get; }

    /// <summary>
    /// 获取或设置表示指定范围段落设置的ParagraphFormat对象。
    /// </summary>
    IWordParagraphFormat? ParagraphFormat { get; set; }

    /// <summary>
    /// 获取表示范围所有列表格式特征的ListFormat对象。
    /// </summary>
    IWordListFormat? ListFormat { get; }

    /// <summary>
    /// 获取表示范围内所有书签的Bookmarks集合。
    /// </summary>
    IWordBookmarks? Bookmarks { get; }

    /// <summary>
    /// 获取一个32位整数，指示创建指定对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取或设置一个值，指示字体或范围是否格式化为粗体。
    /// </summary>
    int Bold { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示范围是否格式化为斜体。
    /// </summary>
    int Italic { get; set; }

    /// <summary>
    /// 获取或设置应用于范围的下划线类型。
    /// </summary>
    WdUnderline Underline { get; set; }

    /// <summary>
    /// 获取或设置字符或指定字符串的强调标记。
    /// </summary>
    WdEmphasisMark EmphasisMark { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示Microsoft Word是否为范围忽略每行的字符数。
    /// </summary>
    bool DisableCharacterSpaceGrid { get; set; }

    /// <summary>
    /// 获取表示范围内跟踪更改的Revisions集合。
    /// </summary>
    IWordRevisions? Revisions { get; }


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
    /// 获取包含指定范围的文章中的字符数。
    /// </summary>
    int StoryLength { get; }

    /// <summary>
    /// 获取或设置指定对象的语言。
    /// </summary>
    WdLanguageID LanguageID { get; set; }

    /// <summary>
    /// 获取包含指定单词或短语同义词、反义词或相关单词信息的SynonymInfo对象。
    /// </summary>
    IWordSynonymInfo? SynonymInfo { get; }

    /// <summary>
    /// 获取表示指定范围内所有超链接的Hyperlinks集合。
    /// </summary>
    IWordHyperlinks? Hyperlinks { get; }

    /// <summary>
    /// 获取表示范围内所有编号段落的ListParagraphs集合。
    /// </summary>
    IWordListParagraphs? ListParagraphs { get; }

    /// <summary>
    /// 获取表示指定范围内所有子文档的Subdocuments集合。
    /// </summary>
    IWordSubdocuments? Subdocuments { get; }

    /// <summary>
    /// 获取或设置一个值，指示是否已对指定范围进行语法检查。
    /// </summary>
    bool GrammarChecked { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否已对指定范围进行拼写检查。
    /// </summary>
    bool SpellingChecked { get; set; }

    /// <summary>
    /// 获取或设置指定范围的突出显示颜色。
    /// </summary>
    WdColorIndex HighlightColorIndex { get; set; }

    /// <summary>
    /// 获取表示范围内所有表格列的Columns集合。
    /// </summary>
    IWordColumns? Columns { get; }

    /// <summary>
    /// 获取表示范围内所有表格行的Rows集合。
    /// </summary>
    IWordRows? Rows { get; }

    /// <summary>
    /// 获取一个值，指示指定范围是否折叠并位于表格的行尾标记处。
    /// </summary>
    bool IsEndOfRowMark { get; }

    /// <summary>
    /// 获取包含指定选择或范围开头的书签编号；如果没有对应的书签，则返回0。
    /// </summary>
    int BookmarkID { get; }

    /// <summary>
    /// 获取在指定范围之前或同一位置开始的最后一个书签的编号。
    /// </summary>
    int PreviousBookmarkID { get; }

    /// <summary>
    /// 获取包含查找操作条件的Find对象。
    /// </summary>
    IWordFind? Find { get; }

    /// <summary>
    /// 获取或设置与指定范围关联的PageSetup对象。
    /// </summary>
    IWordPageSetup? PageSetup { get; set; }

    /// <summary>
    /// 获取表示指定范围内所有Shape对象的ShapeRange集合。
    /// </summary>
    IWordShapeRange? ShapeRange { get; }

    /// <summary>
    /// 获取或设置表示指定范围文本大小写的WdCharacterCase常量。
    /// </summary>
    WdCharacterCase Case { get; set; }

    /// <summary>
    /// 获取表示指定范围可读性统计信息的ReadabilityStatistics集合。
    /// </summary>
    IWordReadabilityStatistics? ReadabilityStatistics { get; }

    /// <summary>
    /// 获取表示范围内语法检查失败的句子的ProofreadingErrors集合。
    /// </summary>
    IWordProofreadingErrors? GrammaticalErrors { get; }

    /// <summary>
    /// 获取表示范围内被标识为拼写错误的单词的ProofreadingErrors集合。
    /// </summary>
    IWordProofreadingErrors? SpellingErrors { get; }

    /// <summary>
    /// 获取或设置启用文本方向功能时范围中文本的方向。
    /// </summary>
    WdTextOrientation Orientation { get; set; }

    /// <summary>
    /// 获取表示文档、范围或选择中所有InlineShape对象的InlineShapes集合。
    /// </summary>
    IWordInlineShapes? InlineShapes { get; }

    /// <summary>
    /// 获取引用下一个故事的范围对象。
    /// </summary>
    IWordRange? NextStoryRange { get; }

    /// <summary>
    /// 获取或设置指定对象的东亚语言。
    /// </summary>
    WdLanguageID LanguageIDFarEast { get; set; }

    /// <summary>
    /// 获取或设置指定对象的其他语言。
    /// </summary>
    WdLanguageID LanguageIDOther { get; set; }

    /// <summary>
    /// 获取或设置指定范围的字符宽度。
    /// </summary>
    WdCharacterWidth CharacterWidth { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否将水平文本设置在垂直文本内。
    /// </summary>
    WdHorizontalInVerticalType HorizontalInVertical { get; set; }

    /// <summary>
    /// 获取或设置Microsoft Word是否在一行中设置两行文本，并指定包围文本的字符（如果有）。
    /// </summary>
    WdTwoLinesInOneType TwoLinesInOne { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示指定范围是否包含组合字符。
    /// </summary>
    bool CombineCharacters { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示Microsoft Word是否已检测到指定文本的语言。
    /// </summary>
    bool LanguageDetected { get; set; }

    /// <summary>
    /// 获取或设置Microsoft Word在当前范围中适应文本的宽度（以当前度量单位表示）。
    /// </summary>
    float FitTextWidth { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示拼写和语法检查器是否忽略指定文本。
    /// </summary>
    int NoProofing { get; set; }

    /// <summary>
    /// 获取表示最外层嵌套级别表格的Tables集合。
    /// </summary>
    IWordTables? TopLevelTables { get; }

    /// <summary>
    /// 获取表示指定对象中HTML脚本集合的Scripts集合。
    /// </summary>
    IOfficeScripts? Scripts { get; }

    /// <summary>
    /// 获取或设置一个值，指示是否显示所有非打印字符。
    /// </summary>
    bool ShowAll { get; set; }

    /// <summary>
    /// 获取与指定范围关联的Document对象。
    /// </summary>
    IWordDocument? Document { get; }

    /// <summary>
    /// 获取表示范围内脚注选项的FootnoteOptions对象。
    /// </summary>
    IWordFootnoteOptions? FootnoteOptions { get; }

    /// <summary>
    /// 获取表示范围内尾注选项的EndnoteOptions对象。
    /// </summary>
    IWordEndnoteOptions? EndnoteOptions { get; }

    /// <summary>
    /// 获取表示范围内所有数学对象的OMaths集合。
    /// </summary>
    IWordOMaths? OMaths { get; }

    /// <summary>
    /// 获取表示用于格式化一个或多个字符的样式的对象。
    /// </summary>
    object CharacterStyle { get; }

    /// <summary>
    /// 获取表示用于格式化段落的样式的对象。
    /// </summary>
    object ParagraphStyle { get; }

    /// <summary>
    /// 获取表示用于格式化项目符号列表或编号列表的样式的对象。
    /// </summary>
    object ListStyle { get; }

    /// <summary>
    /// 获取表示用于格式化表格的样式的对象。
    /// </summary>
    object TableStyle { get; }

    /// <summary>
    /// 获取表示范围内包含的所有内容控件的ContentControls集合。
    /// </summary>
    IWordContentControls? ContentControls { get; }

    /// <summary>
    /// 获取以Microsoft Office Word开放XML格式表示的范围内包含的XML字符串。
    /// </summary>
    string WordOpenXML { get; }

    /// <summary>
    /// 获取表示范围内父内容控件的ContentControl对象。
    /// </summary>
    IWordContentControl? ParentContentControl { get; }

    /// <summary>
    /// 获取表示范围内所有锁定的CoAuthLocks集合。
    /// </summary>
    IWordCoAuthLocks? Locks { get; }

    /// <summary>
    /// 获取表示范围内所有可用更新的CoAuthUpdates集合。
    /// </summary>
    IWordCoAuthUpdates? Updates { get; }

    /// <summary>
    /// 获取包含范围内所有冲突对象的Conflicts集合。
    /// </summary>
    IWordConflicts? Conflicts { get; }

    /// <summary>
    /// 获取一个值，指示屏幕上可见的文本数量。
    /// </summary>
    int TextVisibleOnScreen { get; }

    /// <summary>
    /// 选择指定的对象。
    /// </summary>
    void Select();

    /// <summary>
    /// 设置范围的起始和结束字符位置。
    /// </summary>
    /// <param name="start">范围或选择的起始字符位置。</param>
    /// <param name="end">范围或选择的结束字符位置。</param>
    void SetRange(int start, int end);

    /// <summary>
    /// 将范围折叠到起始或结束位置。
    /// </summary>
    /// <param name="direction">折叠范围或选择的方向。可以是wdCollapseEnd或wdCollapseStart。默认值为wdCollapseStart。</param>
    void Collapse(WdCollapseDirection? direction = null);

    /// <summary>
    /// 在指定范围之前插入指定文本。
    /// </summary>
    /// <param name="text">要插入的文本。</param>
    void InsertBefore(string text);

    /// <summary>
    /// 在范围或选择的末尾插入指定文本。
    /// </summary>
    /// <param name="text">要插入的文本。</param>
    void InsertAfter(string text);

    /// <summary>
    /// 返回表示相对于指定范围的指定单位的新范围对象。
    /// </summary>
    /// <param name="unit">要计数的单位类型。可以是任何WdUnits常量。</param>
    /// <param name="count">要向前移动的单位数。默认值为1。</param>
    /// <returns>新的范围对象。</returns>
    IWordRange? Next(WdUnits? unit = null, int? count = null);

    /// <summary>
    /// 返回相对于指定选择或范围的Range对象。
    /// </summary>
    /// <param name="unit">单位类型。</param>
    /// <param name="count">要向后移动的单位数。默认值为1。</param>
    /// <returns>前一个范围对象。</returns>
    IWordRange? Previous(WdUnits? unit = null, int? count = null);

    /// <summary>
    /// 将指定范围或选择的起始位置移动或扩展到最近指定文本单位的开头。
    /// </summary>
    /// <param name="unit">要移动指定范围或选择起始位置的单位。</param>
    /// <param name="extend">移动类型。</param>
    /// <returns>操作后的字符位置。</returns>
    int? StartOf(WdUnits? unit = null, WdMovementType? extend = null);

    /// <summary>
    /// 将范围或选择的结束字符位置移动或扩展到最近指定文本单位的末尾。
    /// </summary>
    /// <param name="unit">要移动结束字符位置的单位。</param>
    /// <param name="extend">移动类型。</param>
    /// <returns>操作后的字符位置。</returns>
    int? EndOf(WdUnits? unit = null, WdMovementType? extend = null);

    /// <summary>
    /// 将指定的范围或选择折叠到其起始或结束位置，然后将折叠的对象移动指定数量的单位。
    /// </summary>
    /// <param name="unit">折叠范围或选择要移动的单位。</param>
    /// <param name="count">要移动的单元数。</param>
    /// <returns>移动的距离。</returns>
    int? Move(WdUnits? unit = null, int? count = null);

    /// <summary>
    /// 移动指定范围的起始位置。
    /// </summary>
    /// <param name="unit">要移动指定范围或选择起始位置的单位。</param>
    /// <param name="count">要移动的最大单位数。</param>
    /// <returns>移动的距离。</returns>
    int? MoveStart(WdUnits? unit = null, int? count = null);

    /// <summary>
    /// 移动范围的结束字符位置。
    /// </summary>
    /// <param name="unit">要移动结束字符位置的单位。</param>
    /// <param name="count">要移动的单元数。</param>
    /// <returns>移动的距离。</returns>
    int? MoveEnd(WdUnits? unit = null, int? count = null);

    /// <summary>
    /// 在文档中找到任何指定字符时移动指定范围。
    /// </summary>
    /// <param name="characters">一个或多个字符。此参数区分大小写。</param>
    /// <param name="count">要移动的最大字符数。</param>
    /// <returns>移动的距离。</returns>
    int? MoveWhile(string characters, int? count = null);

    /// <summary>
    /// 在文档中找到任何指定字符时移动指定范围的起始位置。
    /// </summary>
    /// <param name="characters">一个或多个字符。此参数区分大小写。</param>
    /// <param name="count">要移动的最大字符数。</param>
    /// <returns>移动的距离。</returns>
    int? MoveStartWhile(string characters, int? count = null);

    /// <summary>
    /// 在文档中找到任何指定字符时移动范围的结束字符位置。
    /// </summary>
    /// <param name="characters">一个或多个字符。此参数区分大小写。</param>
    /// <param name="count">要移动的最大字符数。</param>
    /// <returns>移动的距离。</returns>
    int? MoveEndWhile(string characters, int? count = null);

    /// <summary>
    /// 移动指定范围，直到在文档中找到指定的字符之一。
    /// </summary>
    /// <param name="characters">一个或多个字符。</param>
    /// <param name="count">要移动的最大字符数。</param>
    /// <returns>移动的距离。</returns>
    int? MoveUntil(string characters, int? count = null);

    /// <summary>
    /// 移动指定范围或选择的起始位置，直到在文档中找到指定的字符之一。
    /// </summary>
    /// <param name="characters">一个或多个字符。此参数区分大小写。</param>
    /// <param name="count">要移动的最大字符数。</param>
    /// <returns>移动的距离。</returns>
    int? MoveStartUntil(string characters, int? count = null);

    /// <summary>
    /// 移动范围的结束边界，直到遇到指定字符集中的字符为止，或移动指定次数。
    /// </summary>
    /// <param name="cset">要查找的字符集（字符串）。</param>
    /// <param name="count">移动的单位数量，默认为wdForward（向前移动）。</param>
    /// <returns>返回实际移动的字符数；如果操作失败，则返回null。</returns>
    int? MoveEndUntil(string cset, int? count = null);

    /// <summary>
    /// 从文档中剪切指定的对象并将其放置在剪贴板上。
    /// </summary>
    void Cut();

    /// <summary>
    /// 将指定对象复制到剪贴板。
    /// </summary>
    void Copy();

    /// <summary>
    /// 将剪贴板的内容粘贴到指定范围。
    /// </summary>
    void Paste();

    /// <summary>
    /// 插入分页符、分栏符或分节符。
    /// </summary>
    /// <param name="type">要插入的断点类型。可以是WdBreakType常量之一。</param>
    void InsertBreak(WdBreakType? type = null);

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
    /// 确定应用此方法的选择或范围是否与指定范围在同一故事中。
    /// </summary>
    /// <param name="range">要比较其故事的范围对象。</param>
    /// <returns>如果在同一故事中，则为true；否则为false。</returns>
    bool? InStory(IWordRange range);

    /// <summary>
    /// 确定应用此方法的范围是否包含在指定范围内。
    /// </summary>
    /// <param name="range">要比较的范围对象。</param>
    /// <returns>如果包含在指定范围内，则为true；否则为false。</returns>
    bool? InRange(IWordRange range);

    /// <summary>
    /// 删除指定数量的字符或单词。
    /// </summary>
    /// <param name="unit">要删除的单位。可以是wdCharacter或wdWord。</param>
    /// <param name="count">要删除的单位数。</param>
    /// <returns>删除的数量。</returns>
    int? Delete(WdUnits? unit = null, int? count = null);

    /// <summary>
    /// 展开范围以包含整个故事。
    /// </summary>
    void WholeStory();

    /// <summary>
    /// 展开指定的范围。
    /// </summary>
    /// <param name="unit">要展开范围的单位。可以是WdUnits常量之一。</param>
    /// <returns>扩展后的大小。</returns>
    int? Expand(WdUnits? unit = null);

    /// <summary>
    /// 将指定范围替换为新段落。
    /// </summary>
    void InsertParagraph();

    /// <summary>
    /// 在范围后插入段落标记。
    /// </summary>
    void InsertParagraphAfter();

    /// <summary>
    /// 在范围前插入段落标记。
    /// </summary>
    void InsertParagraphBefore();

    /// <summary>
    /// 在当前范围插入一个特殊符号。
    /// </summary>
    /// <param name="characterNumber">要插入的符号的字符代码。</param>
    /// <param name="font">符号的字体名称，如果为null则使用当前字体。</param>
    /// <param name="Unicode">是否使用Unicode编码，如果为null则使用默认设置。</param>
    /// <param name="Bias">字体选择的倾向性，如果为null则使用默认设置。</param>
    void InsertSymbol(int characterNumber, string? font = null, bool? Unicode = null, WdFontBias? Bias = null);

    /// <summary>
    /// 将选定内容作为图片复制。
    /// </summary>
    void CopyAsPicture();

    /// <summary>
    /// 按升序字母数字顺序对段落或表格行进行排序。
    /// </summary>
    void SortAscending();

    /// <summary>
    /// 按降序字母数字顺序对段落或表格行进行排序。
    /// </summary>
    void SortDescending();

    /// <summary>
    /// 确定应用此方法的范围是否等于指定范围。
    /// </summary>
    /// <param name="range">要比较的范围对象。</param>
    /// <returns>如果相等，则为true；否则为false。</returns>
    bool? IsEqual(IWordRange range);

    /// <summary>
    /// 计算范围内的数学表达式。
    /// </summary>
    /// <returns>计算结果。</returns>
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
    /// 返回引用What参数指定的下一个项目或位置起始位置的Range对象。
    /// </summary>
    /// <param name="what">项目类型。</param>
    /// <returns>新的范围对象。</returns>
    IWordRange? GoToNext(WdGoToItem what);

    /// <summary>
    /// 返回引用What参数指定的前一个项目或位置起始位置的Range对象。
    /// </summary>
    /// <param name="what">项目类型。</param>
    /// <returns>新的范围对象。</returns>
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
    /// 在全局通讯簿列表中查找名称并显示属性对话框。
    /// </summary>
    void LookupNameProperties();

    /// <summary>
    /// 返回基于指定范围内容的统计信息。
    /// </summary>
    /// <param name="statistic">统计类型。</param>
    /// <returns>统计值。</returns>
    int? ComputeStatistics(WdStatistic statistic);

    /// <summary>
    /// 在大纲视图中，将指定范围内的段落移动到下一个可见段落之后或上一个可见段落之前。
    /// </summary>
    /// <param name="direction">移动的方向。</param>
    void Relocate(int direction);

    /// <summary>
    /// 显示同义词库对话框，列出指定范围文本的替代单词选择或同义词。
    /// </summary>
    void CheckSynonyms();

    /// <summary>
    /// 尝试将指定范围中的文本或范围周围的文本与现有的自动图文集条目名称匹配。
    /// </summary>
    void InsertAutoText();

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
    /// 自动格式化范围。
    /// </summary>
    void AutoFormat();

    /// <summary>
    /// 开始对指定范围进行拼写和语法检查。
    /// </summary>
    void CheckGrammar();

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
    /// 将范围移动到下一个子文档。
    /// </summary>
    void NextSubdocument();

    /// <summary>
    /// 将范围或选择移动到上一个子文档。
    /// </summary>
    void PreviousSubdocument();

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
    /// 将单元格或单元格组作为嵌套表格粘贴到选定范围中。
    /// </summary>
    void PasteAsNestedTable();

    /// <summary>
    /// 修改封装样式。
    /// </summary>
    /// <param name="style">封装样式</param>
    /// <param name="symbol">封装符号类型</param>
    /// <param name="enclosedText">被封装的文本，如果为null则使用当前范围的文本</param>
    void ModifyEnclosure(WdEncloseStyle? style, WdEnclosureType symbol, string? enclosedText = null);

    /// <summary>
    /// 为指定范围添加拼音指南。
    /// </summary>
    /// <param name="text">要添加的拼音文本。</param>
    /// <param name="alignment">添加拼音文本的对齐方式。</param>
    /// <param name="raise">从指定范围文本顶部到拼音文本顶部的距离。</param>
    /// <param name="fontSize">拼音文本使用的字体大小。</param>
    /// <param name="fontName">拼音文本使用的字体名称。</param>
    void PhoneticGuide(string text, WdPhoneticGuideAlignmentType alignment = WdPhoneticGuideAlignmentType.wdPhoneticGuideAlignmentCenter, int raise = 0, int fontSize = 0, string fontName = "");

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
    /// 分析指定文本以确定其编写的语言。
    /// </summary>
    void DetectLanguage();

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
    /// 将指定范围从繁体中文转换为简体中文，反之亦然。
    /// </summary>
    /// <param name="direction">转换方向。</param>
    /// <param name="commonTerms">是否将常用表达作为一个整体转换，而不是逐个字符转换。</param>
    /// <param name="useVariants">从简体中文转换为繁体中文时是否使用台湾、香港和澳门的字符变体。</param>
    void TCSCConverter(WdTCSCConverterDirection direction = WdTCSCConverterDirection.wdTCSCConverterDirectionAuto, bool commonTerms = false, bool useVariants = false);

    /// <summary>
    /// 按指定方式粘贴选定的表格单元格并设置其格式。
    /// </summary>
    /// <param name="type">粘贴选定表格单元格时要使用的格式化类型。</param>
    void PasteAndFormat(WdRecoveryType type);

    /// <summary>
    /// 粘贴并格式化Excel表格。
    /// </summary>
    /// <param name="linkedToExcel">是否将粘贴的表格链接到原始Excel文件。</param>
    /// <param name="wordFormatting">是否使用Word文档中的格式设置表格。</param>
    /// <param name="rtf">是否使用RTF格式粘贴Excel表格。</param>
    void PasteExcelTable(bool linkedToExcel, bool wordFormatting, bool rtf);

    /// <summary>
    /// 通过将粘贴的行插入到选定行之间，将粘贴的单元格合并到现有表格中。
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
    /// 将选定范围导出为文档片段以供使用。
    /// </summary>
    /// <param name="fileName">保存文档片段文件的路径和文件名。</param>
    /// <param name="format">文档片段文件的文件格式。</param>
    void ExportFragment(string fileName, WdSaveFormat format);

    /// <summary>
    /// 为编号列表中的一个或多个项目设置列表级别。
    /// </summary>
    /// <param name="level">指示新列表级别的数字。</param>
    void SetListLevel(short level);

    /// <summary>
    /// 插入一个绝对制表符，始终位于相对于页边距或缩进的相同位置。
    /// </summary>
    /// <param name="alignment">制表符的对齐类型。</param>
    /// <param name="relativeTo">制表符是相对于页边距还是段落缩进。</param>
    void InsertAlignmentTab(int alignment, int relativeTo = 0);

    /// <summary>
    /// 将文档片段导入到文档中的指定范围。
    /// </summary>
    /// <param name="fileName">存储文档片段的路径和文件名。</param>
    /// <param name="matchDestination">是否匹配目标格式。</param>
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

}