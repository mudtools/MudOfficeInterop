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
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordSelection : IDisposable
{
    /// <summary>
    /// 获取当前文档归属的<see cref="IWordApplication"/>对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }


    /// <summary>
    /// 获取父对象（通常是 Document）
    /// </summary>
    object? Parent { get; }

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
    /// 获取或设置选择区域的起始位置
    /// </summary>
    int Start { get; set; }

    /// <summary>
    /// 获取或设置选择区域的结束位置
    /// </summary>
    int End { get; set; }


    /// <summary>
    /// 获取关联的文档
    /// </summary>
    IWordDocument? Document { get; }

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
    /// 获取当前选择区域中的子形状范围（如果选择包含形状或形状集合）
    /// </summary>
    IWordShapeRange? ChildShapeRange { get; }

    /// <summary>
    /// 获取当前选择区域中的超链接集合
    /// </summary>
    IWordHyperlinks? Hyperlinks { get; }

    /// <summary>
    /// 指示当前选择是否包含子形状范围
    /// </summary>
    bool HasChildShapeRange { get; }

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
        string? iconFileName, string? iconLabel);

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
    /// 将选择区域向左移动指定的单位数量
    /// </summary>
    /// <param name="unit">移动的单位，默认为单词(wdWord)</param>
    /// <param name="count">移动的单位数量，默认为1</param>
    /// <param name="extend">移动类型，wdMove为移动选择区域，wdExtend为扩展选择区域，默认为wdMove</param>
    /// <returns>移动的字符数，如果失败则返回null</returns>
    int? MoveLeft(WdUnits? unit = WdUnits.wdWord, int? count = 1, WdMovementType? extend = WdMovementType.wdMove);


    /// <summary>
    /// 将选择区域向右移动指定的单位数量
    /// </summary>
    /// <param name="unit">移动的单位，默认为单词(wdWord)</param>
    /// <param name="count">移动的单位数量，默认为1</param>
    /// <param name="extend">移动类型，wdMove为移动选择区域，wdExtend为扩展选择区域，默认为wdMove</param>
    /// <returns>移动的字符数，如果失败则返回null</returns>
    int? MoveRight(WdUnits? unit = WdUnits.wdWord, int? count = 1, WdMovementType? extend = WdMovementType.wdMove);


    /// <summary>
    /// 将选择区域向上移动指定的单位数量
    /// </summary>
    /// <param name="unit">移动的单位，默认为单词(wdWord)</param>
    /// <param name="count">移动的单位数量，默认为1</param>
    /// <param name="extend">移动类型，wdMove为移动选择区域，wdExtend为扩展选择区域，默认为wdMove</param>
    /// <returns>移动的字符数，如果失败则返回null</returns>
    int? MoveUp(WdUnits? unit = WdUnits.wdWord, int? count = 1, WdMovementType? extend = WdMovementType.wdMove);


    /// <summary>
    /// 将选择区域向下移动指定的单位数量
    /// </summary>
    /// <param name="unit">移动的单位，默认为单词(wdWord)</param>
    /// <param name="count">移动的单位数量，默认为1</param>
    /// <param name="extend">移动类型，wdMove为移动选择区域，wdExtend为扩展选择区域，默认为wdMove</param>
    /// <returns>移动的字符数，如果失败则返回null</returns>
    int? MoveDown(WdUnits? unit = WdUnits.wdWord, int? count = 1, WdMovementType? extend = WdMovementType.wdMove);

    /// <summary>
    /// 确定当前选择区域是否在指定的范围区域内
    /// </summary>
    /// <param name="range">要检查的范围区域</param>
    /// <returns>如果当前选择区域在指定范围内则返回true，否则返回false</returns>
    bool? InRange(IWordRange range);

    /// <summary>
    /// 将选择区域收缩到插入点位置
    /// </summary>
    void Shrink();

    /// <summary>
    /// 在表格中的当前行处分割表格
    /// </summary>
    void SplitTable();

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
    /// 将选择区域折叠到起始或结束位置
    /// </summary>
    /// <param name="collapseDirection">折叠方向，null表示折叠到起始位置，wdCollapseEnd表示折叠到结束位置</param>
    void Collapse(WdCollapseDirection? collapseDirection = null);

    /// <summary>
    /// 选择整个文档内容
    /// </summary>
    void WholeStory();

    /// <summary>
    /// 在选择区域插入单元格
    /// </summary>
    /// <param name="ShiftCells">指定如何移动现有单元格，wdInsertCellsShiftRight向右移动单元格，wdInsertCellsShiftDown向下移动单元格</param>
    void InsertCells(WdInsertCells? ShiftCells = null);

    /// <summary>
    /// 获取下一个字段对象
    /// </summary>
    /// <returns>下一个字段对象，如果没有更多字段则返回null</returns>
    IWordField? NextField();

    /// <summary>
    /// 获取上一个字段对象
    /// </summary>
    /// <returns>上一个字段对象，如果没有更多字段则返回null</returns>
    IWordField? PreviousField();

    /// <summary>
    /// 将插入点移动到行首
    /// </summary>
    /// <param name="units">移动单位，默认为单词</param>
    /// <param name="extend">移动时如何扩展选择，wdMove移动不选择，wdExtend选择模式</param>
    /// <returns>移动后的新位置，失败则返回null</returns>
    int? HomeKey(WdUnits units = WdUnits.wdWord, WdMovementType extend = WdMovementType.wdMove);

    /// <summary>
    /// 将插入点移动到行尾
    /// </summary>
    /// <param name="units">移动单位，默认为单词</param>
    /// <param name="extend">移动时如何扩展选择，wdMove移动不选择，wdExtend选择模式</param>
    /// <returns>移动后的新位置，失败则返回null</returns>
    int? EndKey(WdUnits units = WdUnits.wdWord, WdMovementType extend = WdMovementType.wdMove);

    /// <summary>
    /// 跳转到上一个指定类型的项目
    /// </summary>
    /// <param name="What">要查找的项目类型</param>
    /// <returns>找到的范围对象，如果没有找到则返回null</returns>
    IWordRange? GoToPrevious(WdGoToItem What);

    /// <summary>
    /// 跳转到下一个指定类型的项目
    /// </summary>
    /// <param name="What">要查找的项目类型</param>
    /// <returns>找到的范围对象，如果没有找到则返回null</returns>
    IWordRange? GoToNext(WdGoToItem What);

    /// <summary>
    /// 跳转到指定类型的项目
    /// </summary>
    /// <param name="what">要跳转的项目类型，null表示默认为wdGoToPage</param>
    /// <param name="which">跳转方向，null表示wdGoToNext</param>
    /// <param name="count">跳转的次数，null表示1</param>
    /// <param name="name">特定项目的名称，如书签名</param>
    /// <returns>找到的范围对象，如果没有找到则返回null</returns>
    IWordRange? GoTo(WdGoToItem? what = null, WdGoToDirection? which = null, int? count = null, string? name = null);

    /// <summary>
    /// 计算选择区域中的数学表达式
    /// </summary>
    /// <returns>计算结果的浮点数值，如果无法计算则返回null</returns>
    float? Calculate();

    /// <summary>
    /// 比较当前选择区域与指定范围是否相等
    /// </summary>
    /// <param name="Range">要比较的范围对象</param>
    /// <returns>如果两个范围相等则返回true，否则返回false或null</returns>
    bool? IsEqual(IWordRange Range);

    /// <summary>
    /// 按降序对选择区域内容进行排序
    /// </summary>
    void SortDescending();

    /// <summary>
    /// 按升序对选择区域内容进行排序
    /// </summary>
    void SortAscending();

    /// <summary>
    /// 将选择区域内容复制为图片
    /// </summary>
    void CopyAsPicture();

    /// <summary>
    /// 在选择区域插入符号
    /// </summary>
    /// <param name="characterNumber">符号的字符编号</param>
    /// <param name="font">符号所在的字体名称</param>
    /// <param name="unicode">是否使用Unicode编码</param>
    /// <param name="bias">字体偏移量</param>
    void InsertSymbol(int characterNumber, string? font, bool? unicode, WdFontBias? bias);

    /// <summary>
    /// 删除选择区域内容
    /// </summary>
    /// <param name="unit">删除单位</param>
    /// <param name="count">删除数量</param>
    /// <returns>实际删除的单位数</returns>
    int? Delete(WdUnits? unit, int? count);

    /// <summary>
    /// 检查选择区域是否在指定范围内
    /// </summary>
    /// <param name="Range">要检查的范围对象</param>
    /// <returns>如果选择区域在指定范围内则返回true，否则返回false或null</returns>
    bool? InStory(IWordRange Range);

    /// <summary>
    /// 插入文件内容到选择区域
    /// </summary>
    /// <param name="fileName">要插入的文件名</param>
    /// <param name="range">插入位置范围</param>
    /// <param name="ConfirmConversions">是否确认转换</param>
    /// <param name="Link">是否链接文件</param>
    /// <param name="Attachment">是否作为附件插入</param>
    void InsertFile(string fileName, object? range, bool? ConfirmConversions, bool? Link, bool? Attachment);

    /// <summary>
    /// 在选择区域插入分节符
    /// </summary>
    /// <param name="type">分节符类型</param>
    void InsertBreak(WdBreakType? type);

    /// <summary>
    /// 将选择区域的结束位置移动到指定字符集中的第一个字符之前
    /// </summary>
    /// <param name="cset">字符集</param>
    /// <param name="count">移动次数</param>
    /// <returns>实际移动的单位数</returns>
    int? MoveEndUntil(string? cset = null, int? count = null);

    /// <summary>
    /// 将选择区域的起始位置移动到指定字符集中的第一个字符之前
    /// </summary>
    /// <param name="cset">字符集</param>
    /// <param name="count">移动次数</param>
    /// <returns>实际移动的单位数</returns>
    int? MoveStartUntil(string? cset = null, int? count = null);

    /// <summary>
    /// 将选择区域的结束位置移动到指定字符集中的第一个字符之前
    /// </summary>
    /// <param name="cset">字符集</param>
    /// <param name="count">移动次数</param>
    /// <returns>实际移动的单位数</returns>
    int? MoveUntil(string? cset = null, int? count = null);

    /// <summary>
    /// 将选择区域的结束位置移动到指定字符集中的最后一个字符之后
    /// </summary>
    /// <param name="cset">字符集</param>
    /// <param name="count">移动次数</param>
    /// <returns>实际移动的单位数</returns>
    int? MoveEndWhile(string? cset = null, int? count = null);

    /// <summary>
    /// 将选择区域的起始位置移动到指定字符集中的最后一个字符之后
    /// </summary>
    /// <param name="cset">字符集</param>
    /// <param name="count">移动次数</param>
    /// <returns>实际移动的单位数</returns>
    int? MoveStartWhile(string? cset = null, int? count = null);

    /// <summary>
    /// 将选择区域的结束位置移动到指定字符集中的最后一个字符之后
    /// </summary>
    /// <param name="cset">字符集</param>
    /// <param name="count">移动次数</param>
    /// <returns>实际移动的单位数</returns>
    int? MoveWhile(string? cset = null, int? count = null);

    /// <summary>
    /// 将选择区域的结束位置移动指定单位
    /// </summary>
    /// <param name="unit">移动单位</param>
    /// <param name="count">移动数量</param>
    /// <returns>实际移动的单位数</returns>
    int? MoveEnd(WdUnits? unit = null, int? count = null);

    /// <summary>
    /// 将选择区域的起始位置移动指定单位
    /// </summary>
    /// <param name="unit">移动单位</param>
    /// <param name="count">移动数量</param>
    /// <returns>实际移动的单位数</returns>
    int? MoveStart(WdUnits? unit = null, int? count = null);

    /// <summary>
    /// 将选择区域移动指定单位
    /// </summary>
    /// <param name="unit">移动单位</param>
    /// <param name="count">移动数量</param>
    /// <returns>实际移动的单位数</returns>
    int? Move(WdUnits? unit = null, int? count = null);

    /// <summary>
    /// 将选择区域的结束位置移动到指定单位
    /// </summary>
    /// <param name="unit">移动单位</param>
    /// <param name="extend">移动时如何扩展选择，wdMove移动不选择，wdExtend选择模式</param>
    /// <returns>移动后的新位置，失败则返回null</returns>
    int? EndOf(WdUnits? unit = null, WdMovementType? extend = null);

    /// <summary>
    /// 将选择区域的起始位置移动到指定单位
    /// </summary>
    /// <param name="unit">移动单位</param>
    /// <param name="extend">移动时如何扩展选择，wdMove移动不选择，wdExtend选择模式</param>
    /// <returns>移动后的新位置，失败则返回null</returns>
    int? StartOf(WdUnits? unit = null, WdMovementType? extend = null);

    /// <summary>
    /// 获取相对于指定所选内容的上一个范围。
    /// </summary>
    /// <param name="unit">指定移动或扩展所选内容时使用的度量单位，如字符、单词、段落等</param>
    /// <param name="count">指定要移动或扩展的单位数</param>
    /// <returns>上一个范围对象，如果不存在则返回null</returns>
    IWordRange? Previous(WdUnits? unit = null, int? count = null);

    /// <summary>
    /// 获取相对于指定所选内容的下一个范围。
    /// </summary>
    /// <param name="unit">指定移动或扩展所选内容时使用的度量单位，如字符、单词、段落等</param>
    /// <param name="count">指定要移动或扩展的单位数</param>
    /// <returns>下一个范围对象，如果不存在则返回null</returns>
    IWordRange? Next(WdUnits? unit = null, int? count = null);


    /// <summary>
    /// 扩展选择区域
    /// </summary>
    /// <param name="unit">扩展单位</param>
    void Extend(WdUnits? unit = null);
}