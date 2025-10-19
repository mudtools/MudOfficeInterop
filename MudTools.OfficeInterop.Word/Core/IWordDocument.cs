//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Interop.Word;

namespace MudTools.OfficeInterop.Word;


/// <summary>
/// Word 文档接口，用于操作 Word 文档
/// </summary>
public interface IWordDocument : IDisposable
{
    /// <summary>
    /// 获取当前文档归属的<see cref="IWordApplication"/>对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取文档名称
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取文档完整路径
    /// </summary>
    string FullName { get; }

    /// <summary>
    /// 获取或设置文档的加密提供程序
    /// </summary>
    string EncryptionProvider { get; set; }

    /// <summary>
    /// 获取或设置文档内置属性集合
    /// </summary>
    IOfficeDocumentProperties? BuiltInDocumentProperties { get; }

    /// <summary>
    /// 获取或设置文档作者
    /// </summary>
    string Author { get; set; }

    /// <summary>
    /// 获取或设置文档主题
    /// </summary>
    string Subject { get; set; }

    /// <summary>
    /// 获取或设置文档描述
    /// </summary>
    string Description { get; set; }

    /// <summary>
    /// 获取或设置文档关键字
    /// </summary>
    string Keywords { get; set; }

    /// <summary>
    /// 获取或设置文档公司信息
    /// </summary>
    string Company { get; set; }

    /// <summary>
    /// 获取或设置文档标题
    /// </summary>
    string Title { get; set; }

    /// <summary>
    /// 获取文档路径
    /// </summary>
    string Path { get; }

    /// <summary>
    /// 获取文档是否已修改
    /// </summary>
    bool? Saved { get; set; }

    /// <summary>
    /// 获取文档是否已发送路由
    /// </summary>
    bool? Routed { get; }

    /// <summary>
    /// 获取文档是否为主控文档
    /// </summary>
    bool? IsMasterDocument { get; }

    /// <summary>
    /// 获取或设置是否自动断字
    /// </summary>
    bool? AutoHyphenation { get; set; }

    /// <summary>
    /// 获取或设置是否嵌入 TrueType 字体
    /// </summary>
    bool? EmbedTrueTypeFonts { get; set; }

    /// <summary>
    /// 获取或设置是否保存窗体数据
    /// </summary>
    bool? SaveFormsData { get; set; }

    /// <summary>
    /// 获取文档是否为子文档
    /// </summary>
    bool? IsSubdocument { get; }

    /// <summary>
    /// 获取文档保存格式
    /// </summary>
    int? SaveFormat { get; }

    /// <summary>
    /// 获取或设置是否保存子集字体
    /// </summary>
    bool? SaveSubsetFonts { get; set; }

    /// <summary>
    /// 获取或设置是否仅打印窗体数据到预打印的表单上
    /// </summary>
    bool? PrintFormsData { get; set; }

    /// <summary>
    /// 获取或设置是否显示语法错误
    /// </summary>
    bool? ShowGrammaticalErrors { get; set; }

    /// <summary>
    /// 获取或设置是否已完成拼写检查
    /// </summary>
    bool? SpellingChecked { get; set; }

    /// <summary>
    /// 获取或设置是否显示摘要
    /// </summary>
    bool? ShowSummary { get; set; }

    /// <summary>
    /// 获取或设置是否显示拼写错误
    /// </summary>
    bool? ShowSpellingErrors { get; set; }

    /// <summary>
    /// 获取或设置是否已完成语法检查
    /// </summary>
    bool? GrammarChecked { get; set; }

    /// <summary>
    /// 获取或设置是否打印小数宽度字符
    /// </summary>
    bool? PrintFractionalWidths { get; set; }

    /// <summary>
    /// 获取或设置是否在文本上打印 PostScript
    /// </summary>
    bool? PrintPostScriptOverText { get; set; }

    /// <summary>
    /// 获取或设置打开文档时是否更新样式
    /// </summary>
    bool? UpdateStylesOnOpen { get; set; }

    /// <summary>
    /// 获取或设置是否建议以只读方式打开文档
    /// </summary>
    bool? ReadOnlyRecommended { get; set; }

    /// <summary>
    /// 获取或设置是否对大写字母进行断字
    /// </summary>
    bool? HyphenateCaps { get; set; }

    /// <summary>
    /// 获取或设置断字区域宽度（单位：磅）
    /// </summary>
    int? HyphenationZone { get; set; }

    /// <summary>
    /// 获取或设置摘要长度
    /// </summary>
    int? SummaryLength { get; set; }

    /// <summary>
    /// 获取或设置默认制表位宽度（单位：磅）
    /// </summary>
    float? DefaultTabStop { get; set; }

    /// <summary>
    /// 获取或设置连续断字符的最大数量
    /// </summary>
    int? ConsecutiveHyphensLimit { get; set; }

    /// <summary>
    /// 获取或设置文档是否有路由传阅单
    /// </summary>
    bool? HasRoutingSlip { get; set; }

    /// <summary>
    /// 获取或设置文档类型
    /// </summary>
    WdDocumentKind Kind { get; set; }

    /// <summary>
    /// 获取文档是否受保护
    /// </summary>
    bool ReadOnly { get; }

    /// <summary>
    /// 获取文档保护类型
    /// </summary>
    WdProtectionType ProtectionType { get; }

    /// <summary>
    /// 获取文档类型
    /// </summary>
    WdDocumentType Type { get; }

    /// <summary>
    /// 获取文档的命令栏集合
    /// </summary>
    IOfficeCommandBars CommandBars { get; }

    /// <summary>
    /// 获取文档的批注集合
    /// </summary>
    IWordComments Comments { get; }

    /// <summary>
    /// 获取文档的尾注集合
    /// </summary>
    IWordEndnotes Endnotes { get; }

    /// <summary>
    /// 获取文档的脚注集合
    /// </summary>
    IWordFootnotes Footnotes { get; }

    /// <summary>
    /// 获取文档的单词集合
    /// </summary>
    IWordWords? Words { get; }

    /// <summary>
    /// 获取文档的主要内容范围
    /// </summary>
    IWordRange Content { get; }

    /// <summary>
    /// 获取文档的字符集合
    /// </summary>
    IWordCharacters Characters { get; }

    /// <summary>
    /// 获取文档的域集合
    /// </summary>
    IWordFields Fields { get; }

    /// <summary>
    /// 获取文档的窗体域集合
    /// </summary>
    IWordFormFields FormFields { get; }

    /// <summary>
    /// 获取文档的框架集合
    /// </summary>
    IWordFrames Frames { get; }

    /// <summary>
    /// 获取或设置文档的页面设置
    /// </summary>
    IWordPageSetup PageSetup { get; }

    /// <summary>
    /// 获取文档的窗口集合
    /// </summary>
    IWordWindows Windows { get; }

    /// <summary>
    /// 获取文档的信封对象，用于操作文档中的信封相关内容
    /// </summary>
    IWordEnvelope Envelope { get; }

    /// <summary>
    /// 获取或设置文档的背景形状
    /// </summary>
    IWordShape? Background { get; }

    /// <summary>
    /// 获取文档页数
    /// </summary>
    int PageCount { get; }

    /// <summary>
    /// 获取文档字数
    /// </summary>
    int WordCount { get; }

    /// <summary>
    /// 获取文档段落数
    /// </summary>
    int ParagraphCount { get; }

    /// <summary>
    /// 获取文档表格数
    /// </summary>
    int TableCount { get; }

    /// <summary>
    /// 获取文档书签数
    /// </summary>
    int BookmarkCount { get; }

    /// <summary>
    /// 获取父对象（通常是 Application）
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取文档中的内嵌形状集合。
    /// 内嵌形状是嵌入在文本行中的对象，如图片、图表或OLE对象，它们随着文本移动而移动。
    /// </summary>
    IWordInlineShapes? InlineShapes { get; }

    /// <summary>
    /// 获取文档中的浮动形状集合。
    /// 浮动形状是独立于文本流的对象，可以放置在页面上的任意位置，并可以设置文字环绕方式。
    /// </summary>
    IWordShapes? Shapes { get; }

    /// <summary>
    /// 获取活动窗口
    /// </summary>
    IWordWindow ActiveWindow { get; }

    /// <summary>
    /// 获取文档的活动选择区域
    /// </summary>
    IWordSelection Selection { get; }

    /// <summary>
    /// 获取文档范围集合
    /// </summary>
    IWordStoryRanges StoryRanges { get; }

    /// <summary>
    /// 获取文档书签集合
    /// </summary>
    IWordBookmarks Bookmarks { get; }

    /// <summary>
    /// 获取文档表格集合
    /// </summary>
    IWordTables Tables { get; }

    /// <summary>
    /// 获取文档段落集合
    /// </summary>
    IWordParagraphs Paragraphs { get; }

    /// <summary>
    /// 获取文档节集合
    /// </summary>
    IWordSections Sections { get; }

    /// <summary>
    /// 获取文档样式集合
    /// </summary>
    IWordStyles Styles { get; }

    /// <summary>
    /// 获取文档列表模板集合
    /// </summary>
    IWordListTemplates ListTemplates { get; }

    /// <summary>
    /// 获取文档变量集合
    /// </summary>
    IWordVariables Variables { get; }

    /// <summary>
    /// 获取文档自定义属性集合
    /// </summary>
    IWordCustomProperties CustomProperties { get; }

    /// <summary>
    /// 获取或设置文档视图类型
    /// </summary>
    WdViewType ViewType { get; set; }

    /// <summary>
    /// 获取或设置是否显示段落标记
    /// </summary>
    bool ShowParagraphs { get; set; }

    /// <summary>
    /// 获取或设置是否显示隐藏文字
    /// </summary>
    bool ShowHiddenText { get; set; }

    /// <summary>
    /// 获取或设置文档密码
    /// </summary>
    string Password { set; }

    /// <summary>
    /// 获取文档是否设置了密码保护
    /// </summary>
    bool HasPassword { get; }

    /// <summary>
    /// 获取或设置文档写保护密码
    /// </summary>
    string WritePassword { set; }

    /// <summary>
    /// 获取文档的统计信息，如页数、字数、字符数等
    /// </summary>
    /// <param name="Statistic">指定要计算的统计信息类型</param>
    /// <param name="IncludeFootnotesAndEndnotes">是否包含脚注和尾注，默认为 null</param>
    /// <returns>指定统计信息的数值</returns>
    int ComputeStatistics(WdStatistic Statistic, bool? IncludeFootnotesAndEndnotes = null);

    /// <summary>
    /// 激活文档
    /// </summary>
    void Activate();

    /// <summary>
    /// 保存文档
    /// </summary>
    /// <param name="fileName">文件名（可选）</param>
    /// <param name="fileFormat">文件格式（可选）</param>
    void Save(string? fileName = null, WdSaveFormat fileFormat = WdSaveFormat.wdFormatDocumentDefault);

    /// <summary>
    /// 另存为文档
    /// </summary>
    /// <param name="fileName">文件名</param>
    /// <param name="fileFormat">文件格式</param>
    /// <param name="readOnlyRecommended">是否推荐只读打开</param>
    void SaveAs(string fileName, WdSaveFormat fileFormat = WdSaveFormat.wdFormatDocumentDefault, bool readOnlyRecommended = false);

    /// <summary>
    /// 关闭文档
    /// </summary>
    /// <param name="saveChanges">是否保存更改</param>
    void Close(bool saveChanges = true);

    /// <summary>
    /// 关闭当前文档
    /// </summary>
    /// <param name="saveOptions">指定关闭文档时的保存选项</param>
    void Close(WdSaveOptions saveOptions);

    /// <summary>
    /// 打印文档
    /// </summary>
    /// <param name="background">如果为 true，则将文档打印到后台打印机队列</param>
    /// <param name="append">如果为 true，则将指定的文档追加到活动打印机的当前打印作业中</param>
    /// <param name="range">要打印的文档部分</param>
    /// <param name="outputFileName">要将输出发送到的文件的路径和名称</param>
    /// <param name="item">要打印的项目</param>
    /// <param name="copies">要打印的份数</param>
    /// <param name="pages">要打印的页码范围</param>
    /// <param name="pageType">要打印的页面类型（所有页面、奇数页或偶数页）</param>
    /// <param name="printToFile">如果为 true，则将打印输出发送到文件</param>
    /// <param name="collate">如果为 true，则对多份打印进行校对</param>
    /// <param name="manualDuplexPrint">如果为 true，则在手动双面打印机上打印文档</param>
    /// <param name="printZoomColumn">水平打印的页数</param>
    /// <param name="printZoomRow">垂直打印的页数</param>
    /// <param name="printZoomPaperWidth">缩放页的宽度（百分比）</param>
    /// <param name="printZoomPaperHeight">缩放页的高度（百分比）</param>
    void PrintOut(bool? background = null,
        bool? append = null, WdPrintOutRange? range = null,
        string? outputFileName = null,
        WdPrintOutItem? item = null, int? copies = null, string? pages = null,
        WdPrintOutPages? pageType = null, bool? printToFile = null,
        bool? collate = null, bool? manualDuplexPrint = null,
        int? printZoomColumn = null, int? printZoomRow = null,
        int? printZoomPaperWidth = null, int? printZoomPaperHeight = null);

    /// <summary>
    /// 打印文档
    /// </summary>
    /// <param name="copies">打印份数</param>
    /// <param name="pages">打印页码范围</param>
    void PrintOut(int copies, string pages = "");

    /// <summary>
    /// 保护文档
    /// </summary>
    /// <param name="protectionType">保护类型</param>
    /// <param name="password">密码（可选）</param>
    /// <param name="noReset">是否不重置现有保护</param>
    void Protect(WdProtectionType protectionType, string? password = null, bool? noReset = null);

    /// <summary>
    /// 取消文档保护
    /// </summary>
    /// <param name="password">密码（可选）</param>
    void Unprotect(string? password = null);

    /// <summary>
    /// 检查文档保护状态
    /// </summary>
    /// <returns>是否受保护</returns>
    bool IsProtected();

    /// <summary>
    /// 获取指定范围的文本
    /// </summary>
    /// <param name="start">起始位置</param>
    /// <param name="end">结束位置</param>
    /// <returns>文本内容</returns>
    string GetRangeText(int start, int end);

    /// <summary>
    /// 设置指定范围的文本
    /// </summary>
    /// <param name="start">起始位置</param>
    /// <param name="end">结束位置</param>
    /// <param name="text">文本内容</param>
    void SetRangeText(int start, int end, string text);

    /// <summary>
    /// 插入文本到指定位置
    /// </summary>
    /// <param name="position">插入位置</param>
    /// <param name="text">文本内容</param>
    void InsertText(int position, string text);

    /// <summary>
    /// 插入文件内容
    /// </summary>
    /// <param name="fileName">文件路径</param>
    /// <param name="position">插入位置</param>
    void InsertFile(string fileName, int position = -1);

    /// <summary>
    /// 查找并替换文本
    /// </summary>
    /// <param name="findText">查找文本</param>
    /// <param name="replaceText">替换文本</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <param name="matchWholeWord">是否匹配整个单词</param>
    /// <returns>替换次数</returns>
    int FindAndReplace(string findText, string replaceText, bool matchCase = false, bool matchWholeWord = false);

    /// <summary>
    /// 添加书签
    /// </summary>
    /// <param name="name">书签名称</param>
    /// <param name="start">起始位置</param>
    /// <param name="end">结束位置</param>
    /// <returns>书签对象</returns>
    IWordBookmark AddBookmark(string name, int start, int end);

    /// <summary>
    /// 获取书签
    /// </summary>
    /// <param name="name">书签名称</param>
    /// <returns>书签对象</returns>
    IWordBookmark GetBookmark(string name);

    /// <summary>
    /// 删除书签
    /// </summary>
    /// <param name="name">书签名称</param>
    void DeleteBookmark(string name);

    /// <summary>
    /// 添加表格
    /// </summary>
    /// <param name="rows">行数</param>
    /// <param name="columns">列数</param>
    /// <param name="position">插入位置</param>
    /// <returns>表格对象</returns>
    IWordTable AddTable(int rows, int columns, int position = -1);

    /// <summary>
    /// 添加段落
    /// </summary>
    /// <param name="position">插入位置</param>
    /// <param name="text">段落文本</param>
    /// <returns>段落对象</returns>
    IWordParagraph AddParagraph(int position, string text = "");

    /// <summary>
    /// 添加分节符
    /// </summary>
    /// <param name="position">插入位置</param>
    /// <param name="type">分节符类型</param>
    void AddSectionBreak(int position, int type = 2);

    /// <summary>
    /// 添加分页符
    /// </summary>
    /// <param name="position">插入位置</param>
    void AddPageBreak(int position);

    /// <summary>
    /// 添加页眉
    /// </summary>
    /// <param name="text">页眉文本</param>
    /// <param name="primary">是否为主页眉</param>
    void AddHeader(string text, bool primary = true);

    /// <summary>
    /// 添加页脚
    /// </summary>
    /// <param name="text">页脚文本</param>
    /// <param name="primary">是否为主页脚</param>
    void AddFooter(string text, bool primary = true);

    /// <summary>
    /// 设置页边距
    /// </summary>
    /// <param name="top">上边距</param>
    /// <param name="bottom">下边距</param>
    /// <param name="left">左边距</param>
    /// <param name="right">右边距</param>
    void SetMargins(float top, float bottom, float left, float right);

    /// <summary>
    /// 设置页面方向
    /// </summary>
    /// <param name="landscape">是否横向</param>
    void SetPageOrientation(bool landscape = false);

    /// <summary>
    /// 设置页面大小
    /// </summary>
    /// <param name="width">页面宽度</param>
    /// <param name="height">页面高度</param>
    void SetPageSize(float width, float height);

    /// <summary>
    /// 添加变量
    /// </summary>
    /// <param name="name">变量名称</param>
    /// <param name="value">变量值</param>
    /// <returns>变量对象</returns>
    IWordVariable AddVariable(string name, string value);

    /// <summary>
    /// 获取变量
    /// </summary>
    /// <param name="name">变量名称</param>
    /// <returns>变量值</returns>
    string GetVariable(string name);

    /// <summary>
    /// 删除变量
    /// </summary>
    /// <param name="name">变量名称</param>
    void DeleteVariable(string name);

    /// <summary>
    /// 更新所有字段
    /// </summary>
    void UpdateAllFields();

    /// <summary>
    /// 接受所有修订
    /// </summary>
    void AcceptAllRevisions();

    /// <summary>
    /// 拒绝所有修订
    /// </summary>
    void RejectAllRevisions();


    /// <summary>
    /// 导出为 PDF
    /// </summary>
    /// <param name="fileName">PDF 文件路径</param>
    void ExportAsPdf(string fileName);

    /// <summary>
    /// 刷新文档显示
    /// </summary>
    void Refresh();

    /// <summary>
    /// 根据索引获取范围
    /// </summary>
    /// <param name="start">范围开始索引</param>
    /// <param name="end">范围结束索引</param>
    /// <returns>范围对象</returns>
    IWordRange? Range(int? start = null, int? end = null);

    /// <summary>
    /// 根据索引获取范围
    /// </summary>
    /// <param name="start">范围开始索引</param>
    /// <param name="end">范围结束索引</param>
    /// <returns>范围对象</returns>
    IWordRange? this[int start, int end] { get; }

    /// <summary>
    /// 根据书签名称获取范围
    /// </summary>
    /// <param name="bookmarkName">书签名称</param>
    /// <returns>范围对象</returns>
    IWordRange this[string bookmarkName] { get; }
}
