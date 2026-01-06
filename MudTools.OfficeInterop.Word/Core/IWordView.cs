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
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordView : IOfficeObject<IWordView, MsWord.View>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

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
    IWordZoom? Zoom { get; }

    /// <summary>
    /// 获取或设置修订气球的显示位置（边距或内联）。
    /// </summary>
    WdRevisionsBalloonWidthType RevisionsBalloonWidthType { get; set; }

    /// <summary>
    /// 获取或设置修订气球的宽度（当 RevisionsBalloonWidthType 为 wdBalloonWidthSpecified 时有效）。
    /// </summary>
    float RevisionsBalloonWidth { get; set; }

    /// <summary>
    /// 获取或设置页面颜色
    /// </summary>
    WdPageColor PageColor { get; set; }

    /// <summary>
    /// 获取或设置列宽
    /// </summary>
    WdColumnWidth ColumnWidth { get; set; }

    /// <summary>
    /// 获取修订筛选器对象
    /// </summary>
    IWordRevisionsFilter? RevisionsFilter { get; }

    /// <summary>
    /// 获取审阅者对象
    /// </summary>
    IWordReviewers? Reviewers { get; }

    /// <summary>
    /// 获取或设置是否显示其他作者的修订
    /// </summary>
    bool ShowOtherAuthors { get; set; }

    /// <summary>
    /// 获取或设置是否启用冲突模式
    /// </summary>
    bool ConflictMode { get; set; }

    /// <summary>
    /// 获取或设置修订标记模式
    /// </summary>
    WdRevisionsMode MarkupMode { get; set; }

    /// <summary>
    /// 获取或设置是否显示裁切标记
    /// </summary>
    bool ShowCropMarks { get; set; }

    /// <summary>
    /// 获取或设置是否启用平移模式
    /// </summary>
    bool Panning { get; set; }

    /// <summary>
    /// 获取或设置是否高亮显示标记区域
    /// </summary>
    bool ShowMarkupAreaHighlight { get; set; }

    /// <summary>
    /// 获取或设置阅读布局中截断边距的方式
    /// </summary>
    WdReadingLayoutMargin ReadingLayoutTruncateMargins { get; set; }

    /// <summary>
    /// 获取或设置是否在阅读布局中允许编辑
    /// </summary>
    bool ReadingLayoutAllowEditing { get; set; }

    /// <summary>
    /// 获取或设置阅读布局是否允许多页显示
    /// </summary>
    bool ReadingLayoutAllowMultiplePages { get; set; }

    /// <summary>
    /// 获取或设置阅读布局是否使用实际视图
    /// </summary>
    bool ReadingLayoutActualView { get; set; }

    /// <summary>
    /// 获取或设置是否显示背景
    /// </summary>
    bool DisplayBackgrounds { get; set; }

    /// <summary>
    /// 获取或设置可编辑范围的着色模式
    /// </summary>
    int ShadeEditableRanges { get; set; }

    /// <summary>
    /// 获取或设置是否启用阅读布局视图
    /// </summary>
    bool ReadingLayout { get; set; }

    /// <summary>
    /// 获取或设置是否显示修订气球的连接线
    /// </summary>
    bool RevisionsBalloonShowConnectingLines { get; set; }

    /// <summary>
    /// 获取或设置修订气球的边距位置
    /// </summary>
    WdRevisionsBalloonMargin RevisionsBalloonSide { get; set; }

    /// <summary>
    /// 获取或设置是否显示插入和删除标记
    /// </summary>
    bool ShowInsertionsAndDeletions { get; set; }

    /// <summary>
    /// 获取或设置是否显示智能标记
    /// </summary>
    bool DisplaySmartTags { get; set; }

    /// <summary>
    /// 获取或设置是否显示页面边界
    /// </summary>
    bool DisplayPageBoundaries { get; set; }

    /// <summary>
    /// 获取或设置是否显示可选分页符
    /// </summary>
    bool ShowOptionalBreaks { get; set; }

    /// <summary>
    /// 获取或设置导航到的窗口编号
    /// </summary>
    int BrowseToWindow { get; set; }

    /// <summary>
    /// 获取或设置拆分窗口的特殊窗格
    /// </summary>
    WdSpecialPane SplitSpecial { get; set; }

    /// <summary>
    /// 获取或设置字段阴影显示方式
    /// </summary>
    WdFieldShading FieldShading { get; set; }

    /// <summary>
    /// 获取或设置是否显示图片占位符
    /// </summary>
    bool ShowPicturePlaceHolders { get; set; }

    /// <summary>
    /// 获取或设置需要放大的最小字体大小
    /// </summary>
    int EnlargeFontsLessThan { get; set; }

    /// <summary>
    /// 获取或设置是否显示表格网格线
    /// </summary>
    bool TableGridlines { get; set; }

    /// <summary>
    /// 获取或设置是否显示动画
    /// </summary>
    bool ShowAnimation { get; set; }

    /// <summary>
    /// 获取或设置是否将内容调整到窗口大小
    /// </summary>
    bool WrapToWindow { get; set; }

    /// <summary>
    /// 获取或设置是否显示绘图对象
    /// </summary>
    bool ShowDrawings { get; set; }

    /// <summary>
    /// 获取或设置是否显示格式
    /// </summary>
    bool ShowFormat { get; set; }

    /// <summary>
    /// 获取或设置是否只显示首行
    /// </summary>
    bool ShowFirstLineOnly { get; set; }

    /// <summary>
    /// 获取或设置是否启用放大镜功能
    /// </summary>
    bool Magnifier { get; set; }

    /// <summary>
    /// 获取或设置是否显示邮件合并数据视图
    /// </summary>
    bool MailMergeDataView { get; set; }

    /// <summary>
    /// 获取或设置是否启用草稿视图
    /// </summary>
    bool Draft { get; set; }

    /// <summary>
    /// 获取或设置是否启用全屏视图
    /// </summary>
    bool FullScreen { get; set; }

    /// <summary>
    /// 在大纲视图中显示指定级别的标题
    /// </summary>
    /// <param name="level">要显示的标题级别（1-9）</param>
    void ShowHeading(int level);

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