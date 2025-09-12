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
public interface IWordPageSetup : IDisposable
{
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

    /// <summary>
    /// 获取或设置页面方向（0=纵向，1=横向）
    /// </summary>
    WdOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置文本列集合
    /// </summary>
    /// <remarks>
    /// 使用此属性可访问和操作文档或节中的文本列设置。
    /// 文本列允许将页面内容分为多个垂直列，类似于报纸的排版方式。
    /// </remarks>
    IWordTextColumns TextColumns { get; set; }

    /// <summary>
    /// 获取或设置行号设置
    /// </summary>
    IWordLineNumbering LineNumbering { get; set; }

    /// <summary>
    /// 获取或设置奇偶页页眉页脚是否不同（0=相同，1=不同）
    /// </summary>
    int OddAndEvenPagesHeaderFooter { get; set; }

    /// <summary>
    /// 获取或设置首页页眉页脚是否不同（0=相同，1=不同）
    /// </summary>
    int DifferentFirstPageHeaderFooter { get; set; }

    /// <summary>
    /// 获取或设置分节起始方式
    /// </summary>
    WdSectionStart SectionStart { get; set; }

    /// <summary>
    /// 获取或设置页眉到页面顶端的距离
    /// </summary>
    float HeaderDistance { get; set; }

    /// <summary>
    /// 获取或设置页脚到页面底端的距离
    /// </summary>
    float FooterDistance { get; set; }

    /// <summary>
    /// 获取或设置首页纸张托盘
    /// </summary>
    WdPaperTray FirstPageTray { get; set; }

    /// <summary>
    /// 获取或设置其他页面纸张托盘
    /// </summary>
    WdPaperTray OtherPagesTray { get; set; }

    /// <summary>
    /// 获取或设置页面垂直对齐方式
    /// </summary>
    WdVerticalAlignment VerticalAlignment { get; set; }

    /// <summary>
    /// 获取或设置纸张大小
    /// </summary>
    WdPaperSize PaperSize { get; set; }

    /// <summary>
    /// 获取或设置是否在一页上打印两页内容
    /// </summary>
    bool TwoPagesOnOne { get; set; }

    /// <summary>
    /// 获取或设置装订线是否在顶部
    /// </summary>
    bool GutterOnTop { get; set; }
}