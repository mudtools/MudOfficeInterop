namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定矩形的类型及其包含的信息
/// </summary>
[Guid("2C21A8CF-AB68-3F7E-92F9-B745177DF535")]
public enum WdRectangleType
{
    /// <summary>
    /// 表示被文本占用的空间
    /// </summary>
    wdTextRectangle,

    /// <summary>
    /// 表示被形状占用的空间
    /// </summary>
    wdShapeRectangle,

    /// <summary>
    /// 表示被批注框占用的空间
    /// </summary>
    wdMarkupRectangle,

    /// <summary>
    /// 表示被更多（...）指示符占用的空间，当批注有额外文本时该指示符出现在批注框中
    /// </summary>
    wdMarkupRectangleButton,

    /// <summary>
    /// 表示被页面边框占用的空间
    /// </summary>
    wdPageBorderRectangle,

    /// <summary>
    /// 表示与分隔列的行对应的区域
    /// </summary>
    wdLineBetweenColumnRectangle,

    /// <summary>
    /// 表示被选择工具占用的空间，例如表格左上角的表格选择工具或图像的定位点
    /// </summary>
    wdSelection,

    /// <summary>
    /// 不适用
    /// </summary>
    wdSystem,

    /// <summary>
    /// 表示页面上用于显示修订批注框的空间。仅在"打印"对话框中选择"显示标记的文档"打印时才会打印此空间
    /// </summary>
    wdMarkupRectangleArea,

    /// <summary>
    /// 表示在整页阅读视图中阅读文档时被页面导航按钮占用的空间
    /// </summary>
    wdReadingModeNavigation,

    /// <summary>
    /// 表示被用于查找文档中匹配的跟踪移动对的"转到"按钮占用的空间
    /// </summary>
    wdMarkupRectangleMoveMatch,

    /// <summary>
    /// 表示在整页阅读视图中阅读文档时用于翻页的空间
    /// </summary>
    wdReadingModePanningArea,

    /// <summary>
    /// 表示在 Microsoft Office Outlook 中阅读电子邮件时被电子邮件消息导航按钮占用的空间
    /// </summary>
    wdMailNavArea,

    /// <summary>
    /// 表示被内容控件、公式或文档构建基块（文档内控件）占用的空间
    /// </summary>
    wdDocumentControlRectangle
}