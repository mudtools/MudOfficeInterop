//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;


/// <summary>
/// Word 文档接口，用于操作 Word 文档
/// </summary>
public interface IWordDocument : IDisposable, IEnumerable<IWordRange>
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
    bool Saved { get; set; }

    /// <summary>
    /// 获取文档是否受保护
    /// </summary>
    bool ReadOnly { get; }

    /// <summary>
    /// 获取文档保护类型
    /// </summary>
    WdProtectionType ProtectionType { get; }

    /// <summary>
    /// 获取文档页数
    /// </summary>
    int PageCount { get; }

    IWordWords? Words { get; }

    IWordCharacters Characters { get; }

    IWordFields Fields { get; }

    IWordFormFields FormFields { get; }

    IWordFrames Frames { get; }

    IWordPageSetup PageSetup { get; }

    IWordWindows Windows { get; }

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
    /// 获取文档范围
    /// </summary>
    IWordRange Range { get; }

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
    /// 获取或设置文档写保护密码
    /// </summary>
    string WritePassword { set; }

    /// <summary>
    /// 激活文档
    /// </summary>
    void Activate();

    /// <summary>
    /// 保存文档
    /// </summary>
    /// <param name="fileName">文件名（可选）</param>
    /// <param name="fileFormat">文件格式（可选）</param>
    void Save(string fileName = null, WdSaveFormat fileFormat = WdSaveFormat.wdFormatDocumentDefault);

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
    /// 打印文档
    /// </summary>
    /// <param name="copies">打印份数</param>
    /// <param name="pages">打印页码范围</param>
    void PrintOut(int copies = 1, string pages = "");

    /// <summary>
    /// 保护文档
    /// </summary>
    /// <param name="protectionType">保护类型</param>
    /// <param name="password">密码（可选）</param>
    void Protect(WdProtectionType protectionType, string password = null);

    /// <summary>
    /// 取消文档保护
    /// </summary>
    /// <param name="password">密码（可选）</param>
    void Unprotect(string password = null);

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
    /// <param name="index">范围索引</param>
    /// <returns>范围对象</returns>
    IWordRange this[int index] { get; }

    /// <summary>
    /// 根据书签名称获取范围
    /// </summary>
    /// <param name="bookmarkName">书签名称</param>
    /// <returns>范围对象</returns>
    IWordRange this[string bookmarkName] { get; }
}
