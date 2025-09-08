//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 定义对 Microsoft.Office.Interop.Word.View 对象的二次封装接口。
/// 代表 Word 中文档窗口的视图设置。
/// </summary>
public interface IWordView : IDisposable
{
    /// <summary>
    /// 获取创建 View 对象的应用程序对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取 View 对象的父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置视图的类型（例如，普通视图、页面视图）。
    /// 使用此属性更改视图 [[7]]。
    /// </summary>
    WdViewType Type { get; set; }

    /// <summary>
    /// 获取或设置当前视图中要查看的内容（例如，批注、脚注、页眉页脚）。
    /// 使用此属性可以查看批注、尾注、脚注或文档页眉或页脚 [[7]]。
    /// </summary>
    WdSeekView SeekView { get; set; }

    /// <summary>
    /// 获取或设置是否显示段落标记和其他格式符号。
    /// </summary>
    bool ShowParagraphs { get; set; }

    /// <summary>
    /// 获取或设置是否显示可选的连字符。
    /// </summary>
    bool ShowHyphens { get; set; }

    /// <summary>
    /// 获取或设置是否显示隐藏文字。
    /// </summary>
    bool ShowHiddenText { get; set; }

    /// <summary>
    /// 获取或设置是否显示书签。
    /// </summary>
    bool ShowBookmarks { get; set; }

    /// <summary>
    /// 获取或设置是否显示对象锚点。
    /// </summary>
    bool ShowObjectAnchors { get; set; }

    /// <summary>
    /// 获取或设置是否显示文本边界。
    /// </summary>
    bool ShowTextBoundaries { get; set; }

    /// <summary>
    /// 获取或设置是否显示突出显示。
    /// </summary>
    bool ShowHighlight { get; set; }

    /// <summary>
    /// 获取或设置是否显示字段代码。
    /// </summary>
    bool ShowFieldCodes { get; set; }

    /// <summary>
    /// 获取或设置是否显示表格虚框。
    /// </summary>
    bool ShowTabs { get; set; }

    /// <summary>
    /// 获取或设置是否显示空格。
    /// </summary>
    bool ShowSpaces { get; set; }

    /// <summary>
    /// 获取或设置是否显示隐藏文字。
    /// </summary>
    bool ShowAll { get; set; }

    /// <summary>
    /// 获取或设置是否显示主控文档的子文档。
    /// </summary>
    bool ShowMainTextLayer { get; set; }

    /// <summary>
    /// 获取或设置是否显示批注。
    /// </summary>
    bool ShowComments { get; set; }

    /// <summary>
    /// 获取或设置是否显示墨迹注释。
    /// </summary>
    bool ShowInkAnnotations { get; set; }

    /// <summary>
    /// 获取或设置是否显示 XML 结构。
    /// </summary>
    int ShowXMLMarkup { get; set; }

    /// <summary>
    /// 获取或设置是否显示格式更改。
    /// </summary>
    bool ShowFormatChanges { get; set; }

    /// <summary>
    /// 获取或设置是否显示插入内容修订。
    /// </summary>
    bool ShowRevisionsAndComments { get; set; }

    /// <summary>
    /// 获取或设置修订和批注的显示模式。
    /// </summary>
    WdRevisionsView RevisionsView { get; set; }

    /// <summary>
    /// 获取或设置修订和批注的显示状态。
    /// </summary>
    WdRevisionsMode RevisionsMode { get; set; }

    /// <summary>
    /// 获取与视图关联的缩放对象。
    /// </summary>
    IWordZoom Zoom { get; }

    /// <summary>
    /// 获取或设置修订气球的显示位置（边距或内联）。
    /// </summary>
    WdRevisionsBalloonWidthType RevisionsBalloonWidthType { get; set; }

    /// <summary>
    /// 获取或设置修订气球的宽度（当 RevisionsBalloonWidthType 为 wdBalloonWidthSpecified 时有效）。
    /// </summary>
    float RevisionsBalloonWidth { get; set; }

    // --- 方法封装 ---

    /// <summary>
    /// 展开指定范围的大纲级别
    /// </summary>
    /// <param name="rang">需要展开大纲的文档范围</param>
    void ExpandOutline(IWordRange rang);

    /// <summary>
    /// 折叠指定范围的大纲级别
    /// </summary>
    /// <param name="rang">需要折叠大纲的文档范围</param>
    void CollapseOutline(IWordRange rang);

    /// <summary>
    /// 折叠文档中的所有标题。
    /// </summary>
    void CollapseAllHeadings();

    /// <summary>
    /// 展开文档中的所有标题。
    /// </summary>
    void ExpandAllHeadings();

    /// <summary>
    /// 如果选定内容位于标头中，则移动到当前部分中的下一个标头或下一节中的第一个标头。
    /// 如果选定内容位于页脚中，则移动到下一页脚。
    /// </summary>
    void NextHeaderFooter();

    /// <summary>
    /// 如果选定内容位于标头中，则移动到当前节中的上一个标头或上一节中的最后一个标头。
    /// 如果选定内容位于页脚中，则移动到上一页脚。
    /// </summary>
    void PreviousHeaderFooter();

    /// <summary>
    /// 在显示所有文本（标题和正文文本）和仅显示标题之间切换。
    /// </summary>
    void ShowAllHeadings();

}