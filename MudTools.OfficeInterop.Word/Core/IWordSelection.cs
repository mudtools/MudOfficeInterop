//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word Selection 接口，用于操作 Word 文档中的选择区域
/// </summary>
public interface IWordSelection : IDisposable
{
    /// <summary>
    /// 获取当前文档归属的<see cref="IWordApplication"/>对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取选择区域的文本内容
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取一个值，指示当前选择区域是否处于活动状态
    /// </summary>
    bool Active { get; }

    /// <summary>
    /// 获取一个值，指示插入点是否位于行尾
    /// </summary>
    bool IPAtEndOfLine { get; }

    /// <summary>
    /// 获取或设置一个值，指示选择区域的起始位置是否处于活动状态
    /// </summary>
    bool StartIsActive { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否处于列选择模式
    /// </summary>
    bool ColumnSelectMode { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否处于扩展模式
    /// </summary>
    bool ExtendMode { get; set; }

    /// <summary>
    /// 获取选择区域的类型
    /// </summary>
    WdSelectionType Type { get; }

    /// <summary>
    /// 获取选择区域所在的故事类型（如主文本、页眉、页脚等）
    /// </summary>
    WdStoryType StoryType { get; }

    /// <summary>
    /// 获取选择区域的标志
    /// </summary>
    WdSelectionFlags Flags { get; }

    /// <summary>
    /// 获取当前故事的长度
    /// </summary>
    int StoryLength { get; }

    /// <summary>
    /// 获取或设置选择区域的语言ID
    /// </summary>
    WdLanguageID LanguageID { get; set; }

    /// <summary>
    /// 获取或设置选择区域的远东语言ID
    /// </summary>
    WdLanguageID LanguageIDFarEast { get; set; }

    /// <summary>
    /// 获取或设置选择区域的其他语言ID
    /// </summary>
    WdLanguageID LanguageIDOther { get; set; }

    /// <summary>
    /// 获取一个值，指示选择区域是否位于行末标记处
    /// </summary>
    bool IsEndOfRowMark { get; }

    /// <summary>
    /// 获取与选择区域关联的书签ID
    /// </summary>
    int BookmarkID { get; }

    /// <summary>
    /// 获取与选择区域前一个位置关联的书签ID
    /// </summary>
    int PreviousBookmarkID { get; }

    /// <summary>
    /// 获取选择区域的类型
    /// </summary>
    WdBuiltinStyle Style { get; }

    /// <summary>
    /// 获取或设置选择区域的起始位置
    /// </summary>
    int Start { get; set; }

    /// <summary>
    /// 获取或设置选择区域的结束位置
    /// </summary>
    int End { get; set; }

    /// <summary>
    /// 获取选择区域的长度
    /// </summary>
    int Length { get; }

    /// <summary>
    /// 获取父对象（通常是 Document）
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取关联的文档
    /// </summary>
    IWordDocument? Document { get; }

    /// <summary>
    /// 获取或设置字体名称
    /// </summary>
    string FontName { get; set; }

    /// <summary>
    /// 获取或设置字体大小
    /// </summary>
    float FontSize { get; set; }

    /// <summary>
    /// 获取或设置是否加粗
    /// </summary>
    bool Bold { get; set; }

    /// <summary>
    /// 获取或设置是否斜体
    /// </summary>
    bool Italic { get; set; }

    /// <summary>
    /// 获取或设置下划线类型
    /// </summary>
    int Underline { get; set; }

    /// <summary>
    /// 获取或设置文字颜色
    /// </summary>
    WdColor FontColor { get; set; }

    /// <summary>
    /// 获取或设置段落对齐方式
    /// </summary>
    int Alignment { get; set; }

    /// <summary>
    /// 获取或设置行距
    /// </summary>
    float LineSpacing { get; set; }

    /// <summary>
    /// 获取或设置段前间距
    /// </summary>
    float SpaceBefore { get; set; }

    /// <summary>
    /// 获取或设置段后间距
    /// </summary>
    float SpaceAfter { get; set; }

    /// <summary>
    /// 获取或设置首行缩进
    /// </summary>
    float FirstLineIndent { get; set; }

    /// <summary>
    /// 获取或设置选择区域的文本方向
    /// </summary>
    WdTextOrientation Orientation { get; set; }

    /// <summary>
    /// 获取选择区域的字体格式设置
    /// </summary>
    IWordFont? Font { get; }

    /// <summary>
    /// 获取选择区域中的形状范围集合
    /// </summary>
    IWordShapeRange? ShapeRange { get; }

    /// <summary>
    /// 获取选择区域中的内嵌形状集合
    /// </summary>
    IWordInlineShapes? InlineShapes { get; }

    /// <summary>
    /// 获取选择区域中的段落集合
    /// </summary>
    IWordParagraphs? Paragraphs { get; }

    /// <summary>
    /// 获取选择区域的边框设置
    /// </summary>
    IWordBorders? Borders { get; }

    /// <summary>
    /// 获取选择区域的底纹格式
    /// </summary>
    IWordShading? Shading { get; }

    /// <summary>
    /// 获取选择区域中的字段集合
    /// </summary>
    IWordFields? Fields { get; }

    /// <summary>
    /// 获取选择区域中的窗体字段集合
    /// </summary>
    IWordFormFields? FormFields { get; }

    /// <summary>
    /// 获取选择区域中的框架集合
    /// </summary>
    IWordFrames? Frames { get; }

    /// <summary>
    /// 获取选择区域的段落格式设置
    /// </summary>
    IWordParagraphFormat? ParagraphFormat { get; }

    /// <summary>
    /// 获取选择区域的页面设置
    /// </summary>
    IWordPageSetup? PageSetup { get; }

    /// <summary>
    /// 获取选择区域中的书签集合
    /// </summary>
    IWordBookmarks? Bookmarks { get; }

    /// <summary>
    /// 获取选择区域中的节集合
    /// </summary>
    IWordSections? Sections { get; }

    /// <summary>
    /// 获取选择区域中的单元格集合
    /// </summary>
    IWordCells? Cells { get; }

    /// <summary>
    /// 获取选择区域中的列集合
    /// </summary>
    IWordColumns? Columns { get; }

    /// <summary>
    /// 获取选择区域中的行集合
    /// </summary>
    IWordRows? Rows { get; }

    /// <summary>
    /// 获取选择区域中的页眉/页脚设置
    /// </summary>
    IWordHeaderFooter? HeaderFooter { get; }

    /// <summary>
    /// 获取选择区域中的批注集合
    /// </summary>
    IWordComments? Comments { get; }

    /// <summary>
    /// 获取选择区域中的尾注集合
    /// </summary>
    IWordEndnotes? Endnotes { get; }

    /// <summary>
    /// 获取选择区域中的脚注集合
    /// </summary>
    IWordFootnotes? Footnotes { get; }

    /// <summary>
    /// 获取选择区域中的字符集合
    /// </summary>
    IWordCharacters? Characters { get; }

    /// <summary>
    /// 获取选择区域中的句子集合
    /// </summary>
    IWordSentences? Sentences { get; }

    /// <summary>
    /// 获取选择区域中的表格集合
    /// </summary>
    IWordTables? Tables { get; }

    /// <summary>
    /// 获取选择区域中的单词集合
    /// </summary>
    IWordWords? Words { get; }

    /// <summary>
    /// 获取选择区域的带格式文本
    /// </summary>
    IWordRange? FormattedText { get; }

    /// <summary>
    /// 获取查找对象
    /// </summary>
    IWordFind? Find { get; }

    /// <summary>
    /// 获取范围对象
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 激活选择区域
    /// </summary>
    void Activate();

    /// <summary>
    /// 复制选择区域内容
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切选择区域内容
    /// </summary>
    void Cut();

    /// <summary>
    /// 粘贴内容到选择区域
    /// </summary>
    void Paste();

    /// <summary>
    /// 删除选择区域内容
    /// </summary>
    void Delete();

    /// <summary>
    /// 清除格式
    /// </summary>
    void ClearFormatting();

    /// <summary>
    /// 插入文本
    /// </summary>
    /// <param name="text">要插入的文本</param>
    void InsertText(string text);

    /// <summary>
    /// 插入段落
    /// </summary>
    void InsertParagraph();

    /// <summary>
    /// 在所选内容或指定范围之后插入新段落。
    /// </summary>
    void InsertParagraphAfter();

    /// <summary>
    /// 在所选内容或指定范围之前插入新段落。
    /// </summary>
    void InsertParagraphBefore();

    /// <summary>
    /// 在所选内容或指定范围之前插入文本。
    /// </summary>
    /// <param name="text">要插入的文本</param>
    void InsertBefore(string text);

    /// <summary>
    /// 在所选内容或指定范围之后插入文本。
    /// </summary>
    /// <param name="text">要插入的文本</param>
    void InsertAfter(string text);

    /// <summary>
    /// 在所选内容后插入下一页分节符。
    /// </summary>
    void InsertNewPage();

    /// <summary>
    /// 获取相对于指定所选内容的下一个范围。
    /// </summary>
    /// <param name="unit">指定移动或扩展所选内容时使用的度量单位，如字符、单词、段落等</param>
    /// <param name="count">指定要移动或扩展的单位数</param>
    /// <returns>下一个范围对象，如果不存在则返回null</returns>
    IWordRange? Next(WdUnits unit, int count);

    /// <summary>
    /// 获取相对于指定所选内容的上一个范围。
    /// </summary>
    /// <param name="unit">指定移动或扩展所选内容时使用的度量单位，如字符、单词、段落等</param>
    /// <param name="count">指定要移动或扩展的单位数</param>
    /// <returns>上一个范围对象，如果不存在则返回null</returns>
    IWordRange? Previous(WdUnits unit, int count);

    /// <summary>
    /// 粘贴Excel表格到文档中。
    /// </summary>
    /// <param name="linkedToExcel">如果为true，则粘贴的表格链接到Excel文件</param>
    /// <param name="wordFormatting">如果为true，则使用Word格式</param>
    /// <param name="RTF">如果为true，则以RTF格式粘贴</param>
    void PasteExcelTable(bool linkedToExcel, bool wordFormatting, bool RTF);

    /// <summary>
    /// 粘贴剪贴板中的内容并保留原格式。
    /// </summary>
    void PasteFormat();

    /// <summary>
    /// 将剪贴板中的内容作为嵌套表格粘贴。
    /// </summary>
    void PasteAsNestedTable();

    /// <summary>
    /// 以特殊方式粘贴剪贴板中的内容。
    /// </summary>
    /// <param name="iconIndex">指定要使用的图标索引</param>
    /// <param name="link">指定是否创建链接</param>
    /// <param name="placement">指定OLE对象在文档中的放置方式</param>
    /// <param name="displayAsIcon">指定是否将粘贴的内容显示为图标</param>
    /// <param name="dataType">指定粘贴数据的类型</param>
    /// <param name="iconFileName">指定图标文件名</param>
    /// <param name="iconLabel">指定图标标签</param>
    void PasteSpecial(bool? iconIndex,
        bool? link, WdOLEPlacement? placement,
        bool? displayAsIcon, WdPasteDataType? dataType,
        bool? iconFileName, bool? iconLabel);

    /// <summary>
    /// 将剪贴板内容作为表格追加到现有表格中。
    /// </summary>
    void PasteAppendTable();

    /// <summary>
    /// 粘贴剪贴板中的内容并按照指定格式恢复。
    /// </summary>
    /// <param name="Type">指定恢复格式的类型</param>
    void PasteAndFormat(WdRecoveryType Type);

    /// <summary>
    /// 插入换行符
    /// </summary>
    void InsertLineBreak();

    /// <summary>
    /// 插入分页符
    /// </summary>
    void InsertPageBreak();

    /// <summary>
    /// 插入表格
    /// </summary>
    /// <param name="rows">行数</param>
    /// <param name="columns">列数</param>
    /// <returns>表格对象</returns>
    IWordTable InsertTable(int rows, int columns);

    /// <summary>
    /// 向前移动选择区域
    /// </summary>
    /// <param name="unit">移动单位</param>
    /// <param name="count">移动数量</param>
    /// <returns>实际移动的单位数</returns>
    int MoveLeft(int unit = 1, int count = 1);

    /// <summary>
    /// 向后移动选择区域
    /// </summary>
    /// <param name="unit">移动单位</param>
    /// <param name="count">移动数量</param>
    /// <returns>实际移动的单位数</returns>
    int MoveRight(int unit = 1, int count = 1);

    /// <summary>
    /// 向上移动选择区域
    /// </summary>
    /// <param name="unit">移动单位</param>
    /// <param name="count">移动数量</param>
    /// <returns>实际移动的单位数</returns>
    int MoveUp(int unit = 1, int count = 1);

    /// <summary>
    /// 向下移动选择区域
    /// </summary>
    /// <param name="unit">移动单位</param>
    /// <param name="count">移动数量</param>
    /// <returns>实际移动的单位数</returns>
    int MoveDown(int unit = 1, int count = 1);

    /// <summary>
    /// 确定当前选择区域是否在指定的范围区域内
    /// </summary>
    /// <param name="range">要检查的范围区域</param>
    /// <returns>如果当前选择区域在指定范围内则返回true，否则返回false</returns>
    bool InRange(IWordRange range);

    /// <summary>
    /// 将选择区域收缩到插入点位置
    /// </summary>
    void Shrink();

    /// <summary>
    /// 在表格中的当前行处分割表格
    /// </summary>
    void SplitTable();

    /// <summary>
    /// 将选择区域移动到指定单位的开始位置
    /// </summary>
    /// <param name="unit">移动的单位类型</param>
    /// <param name="extend">移动类型，可以是移动或扩展选择</param>
    /// <returns>移动的位置数，如果操作失败则返回null</returns>
    int? StartOf(WdUnits? unit, WdMovementType? extend);

    /// <summary>
    /// 重新设置选择区域的起始和结束位置
    /// </summary>
    /// <param name="start">新的起始位置</param>
    /// <param name="end">新的结束位置</param>
    void SetRange(int start, int end);

    /// <summary>
    /// 选择当前区域
    /// </summary>
    void Select();

    /// <summary>
    /// 选择当前单元格
    /// </summary>
    void SelectCell();

    /// <summary>
    /// 选择当前列
    /// </summary>
    void SelectColumn();

    /// <summary>
    /// 选择具有当前对齐方式的文本
    /// </summary>
    void SelectCurrentAlignment();

    /// <summary>
    /// 选择具有当前颜色格式的文本
    /// </summary>
    void SelectCurrentColor();

    /// <summary>
    /// 选择具有当前字体格式的文本
    /// </summary>
    void SelectCurrentFont();

    /// <summary>
    /// 选择具有当前缩进格式的文本
    /// </summary>
    void SelectCurrentIndent();

    /// <summary>
    /// 选择具有当前间距格式的文本
    /// </summary>
    void SelectCurrentSpacing();

    /// <summary>
    /// 选择具有当前制表符设置的文本
    /// </summary>
    void SelectCurrentTabs();

    /// <summary>
    /// 选择当前行
    /// </summary>
    void SelectRow();

    /// <summary>
    /// 全选内容
    /// </summary>
    void SelectAll();

    /// <summary>
    /// 取消选择
    /// </summary>
    void Collapse();

    /// <summary>
    /// 扩展选择区域
    /// </summary>
    /// <param name="unit">扩展单位</param>
    /// <param name="count">扩展数量</param>
    void Extend(int unit = 1, int count = 1);

    /// <summary>
    /// 查找并替换文本
    /// </summary>
    /// <param name="findText">查找文本</param>
    /// <param name="replaceText">替换文本</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <param name="matchWholeWord">是否匹配整个单词</param>
    /// <returns>是否找到并替换</returns>
    bool FindAndReplace(string findText, string replaceText, bool matchCase = false, bool matchWholeWord = false);

    /// <summary>
    /// 设置字符格式
    /// </summary>
    /// <param name="fontName">字体名称</param>
    /// <param name="fontSize">字体大小</param>
    /// <param name="bold">是否加粗</param>
    /// <param name="italic">是否斜体</param>
    /// <param name="underline">下划线类型</param>
    /// <param name="color">文字颜色</param>
    void SetFont(string fontName = null, float fontSize = 0, bool bold = false, bool italic = false, int underline = 0, int color = 0);

    /// <summary>
    /// 设置段落格式
    /// </summary>
    /// <param name="alignment">对齐方式</param>
    /// <param name="lineSpacing">行距</param>
    /// <param name="spaceBefore">段前间距</param>
    /// <param name="spaceAfter">段后间距</param>
    /// <param name="firstLineIndent">首行缩进</param>
    void SetParagraph(int alignment = 0, float lineSpacing = 0, float spaceBefore = 0, float spaceAfter = 0, float firstLineIndent = 0);

    /// <summary>
    /// 获取选择区域的书签
    /// </summary>
    /// <param name="name">书签名称</param>
    /// <returns>书签对象</returns>
    IWordBookmark GetBookmark(string name);

    /// <summary>
    /// 添加书签
    /// </summary>
    /// <param name="name">书签名称</param>
    /// <returns>书签对象</returns>
    IWordBookmark AddBookmark(string name);

    /// <summary>
    /// 获取选择区域的超链接
    /// </summary>
    /// <param name="address">超链接地址</param>
    /// <returns>超链接对象</returns>
    IWordHyperlink AddHyperlink(string address);

    /// <summary>
    /// 刷新显示
    /// </summary>
    void Refresh();
}