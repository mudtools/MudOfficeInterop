//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;


/// <summary>
/// Word 文档页面设置接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordPageSetup : IDisposable
{
    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 Application 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    #region 页面尺寸和边距设置

    /// <summary>
    /// 获取或设置上边距
    /// </summary>
    float TopMargin { get; set; }

    /// <summary>
    /// 获取或设置下边距
    /// </summary>
    float BottomMargin { get; set; }

    /// <summary>
    /// 获取或设置左边距
    /// </summary>
    float LeftMargin { get; set; }

    /// <summary>
    /// 获取或设置右边距
    /// </summary>
    float RightMargin { get; set; }

    /// <summary>
    /// 获取或设置页面宽度
    /// </summary>
    float PageWidth { get; set; }

    /// <summary>
    /// 获取或设置页面高度
    /// </summary>
    float PageHeight { get; set; }

    #endregion

    #region 页面方向和布局设置

    /// <summary>
    /// 获取或设置页面方向（0=纵向，1=横向）
    /// </summary>
    WdOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置页面垂直对齐方式
    /// </summary>
    WdVerticalAlignment VerticalAlignment { get; set; }

    /// <summary>
    /// 获取或设置页面布局模式
    /// </summary>
    WdLayoutMode LayoutMode { get; set; }

    #endregion

    #region 装订线设置

    /// <summary>
    /// 获取或设置装订线宽度
    /// </summary>
    float Gutter { get; set; }

    /// <summary>
    /// 获取或设置装订线样式
    /// </summary>
    WdGutterStyleOld GutterStyle { get; set; }

    /// <summary>
    /// 获取或设置装订线位置
    /// </summary>
    WdGutterStyle GutterPos { get; set; }

    /// <summary>
    /// 获取或设置装订线是否在顶部
    /// </summary>
    bool GutterOnTop { get; set; }

    #endregion

    #region 文本列和行号设置

    /// <summary>
    /// 获取或设置文本列集合
    /// </summary>
    /// <remarks>
    /// 使用此属性可访问和操作文档或节中的文本列设置。
    /// 文本列允许将页面内容分为多个垂直列，类似于报纸的排版方式。
    /// </remarks>
    IWordTextColumns? TextColumns { get; set; }

    /// <summary>
    /// 获取或设置行号设置
    /// </summary>
    IWordLineNumbering? LineNumbering { get; set; }

    /// <summary>
    /// 获取或设置每行字符数
    /// </summary>
    float CharsLine { get; set; }

    /// <summary>
    /// 获取或设置每页行数
    /// </summary>
    float LinesPage { get; set; }

    #endregion

    #region 页眉页脚设置

    /// <summary>
    /// 获取或设置奇偶页页眉页脚是否不同（0=相同，1=不同）
    /// </summary>
    int OddAndEvenPagesHeaderFooter { get; set; }

    /// <summary>
    /// 获取或设置首页页眉页脚是否不同（0=相同，1=不同）
    /// </summary>
    int DifferentFirstPageHeaderFooter { get; set; }

    /// <summary>
    /// 获取或设置页眉到页面顶端的距离
    /// </summary>
    float HeaderDistance { get; set; }

    /// <summary>
    /// 获取或设置页脚到页面底端的距离
    /// </summary>
    float FooterDistance { get; set; }

    #endregion

    #region 章节和分节设置

    /// <summary>
    /// 获取或设置章节文本方向
    /// </summary>
    WdSectionDirection SectionDirection { get; set; }

    /// <summary>
    /// 获取或设置分节起始方式
    /// </summary>
    WdSectionStart SectionStart { get; set; }

    #endregion

    #region 书籍折页打印设置

    /// <summary>
    /// 获取或设置书籍折页反向打印设置
    /// </summary>
    bool BookFoldRevPrinting { get; set; }

    /// <summary>
    /// 获取或设置书籍折页打印的纸张数量
    /// </summary>
    int BookFoldPrintingSheets { get; set; }

    /// <summary>
    /// 获取或设置是否启用书籍折页打印模式
    /// </summary>
    bool BookFoldPrinting { get; set; }

    #endregion

    #region 纸张和打印设置

    /// <summary>
    /// 获取或设置首页纸张托盘
    /// </summary>
    WdPaperTray FirstPageTray { get; set; }

    /// <summary>
    /// 获取或设置其他页面纸张托盘
    /// </summary>
    WdPaperTray OtherPagesTray { get; set; }

    /// <summary>
    /// 获取或设置纸张大小
    /// </summary>
    WdPaperSize PaperSize { get; set; }

    /// <summary>
    /// 获取或设置是否在一页上打印两页内容
    /// </summary>
    bool TwoPagesOnOne { get; set; }

    #endregion

    #region 其他设置

    /// <summary>
    /// 获取或设置是否抑制脚注显示（0=不抑制，1=抑制）
    /// </summary>
    int SuppressEndnotes { get; set; }

    /// <summary>
    /// 获取或设置是否显示网格线
    /// </summary>
    bool ShowGrid { get; set; }

    #endregion

    /// <summary>
    /// 切换页面方向为纵向（如果当前是横向）或横向（如果当前是纵向）
    /// </summary>
    void TogglePortrait();

    /// <summary>
    /// 将当前页面设置应用为模板的默认设置
    /// </summary>
    void SetAsTemplateDefault();
}