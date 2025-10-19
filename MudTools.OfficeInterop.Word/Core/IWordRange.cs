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
public interface IWordRange : IDisposable
{
    #region 基本属性 (Basic Properties)

    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
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
    /// 获取或设置区域的样式。
    /// </summary>
    object? Style { get; set; }

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
    bool NoProofing { get; set; }

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

    /// <summary>
    /// 将另一个范围的内容复制到此区域 (通过 FormattedText)。
    /// </summary>
    /// <param name="source">源范围。</param>
    void CopyAsText(IWordRange source);

    /// <summary>
    /// 获取有关区域的信息。
    /// </summary>
    /// <param name="type">信息类型。</param>
    /// <returns>相关信息。</returns>
    object Information(WdInformation type);

    /// <summary>
    /// 查找并替换文本。
    /// </summary>
    /// <param name="findText">要查找的文本。</param>
    /// <param name="replaceWith">替换文本。</param>
    /// <param name="replace">替换操作类型。</param>
    /// <returns>是否找到并替换。</returns>
    bool FindAndReplace(object findText, object replaceWith, MsWord.WdReplace replace);

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

    /// <summary>
    /// 将区域转换为书签。
    /// </summary>
    /// <param name="name">书签名称。</param>
    /// <returns>创建的书签。</returns>
    IWordBookmark? Bookmark(string name);

    /// <summary>
    /// 将区域转换为超链接。
    /// </summary>
    /// <param name="address">链接地址。</param>
    /// <param name="subAddress">子地址。</param>
    /// <param name="screenTip">屏幕提示。</param>
    /// <param name="textToDisplay">显示文本。</param>
    /// <param name="target">目标框架。</param>
    /// <returns>创建的超链接。</returns>
    IWordHyperlink? Hyperlink(object address, object subAddress, object screenTip, object textToDisplay, object target);

    /// <summary>
    /// 合并单元格（如果区域在表格内）。
    /// </summary>
    void CellsMerge();

    /// <summary>
    /// 拆分表格单元格（如果区域在表格内）。
    /// </summary>
    /// <param name="numRows">行数。</param>
    /// <param name="numColumns">列数。</param>
    /// <param name="mergeBeforeSplit">拆分前是否合并。</param>
    void CellsSplit(int numRows, int numColumns, bool mergeBeforeSplit);

    /// <summary>
    /// 排序段落。
    /// </summary>
    /// <param name="excludeHeader">是否排除标题。</param>
    /// <param name="fieldNumber">字段号。</param>
    /// <param name="sortFieldType">排序字段类型。</param>
    /// <param name="ascending">是否升序。</param>
    void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object ascending);

    /// <summary>
    /// 将区域另存为文本文件。
    /// </summary>
    /// <param name="fileName">文件名。</param>
    void SaveAsText(string fileName);

    #endregion
}