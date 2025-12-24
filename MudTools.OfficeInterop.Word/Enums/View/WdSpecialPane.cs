//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定要在活动窗口窗格中显示的项
/// </summary>
public enum WdSpecialPane
{
    /// <summary>
    /// 无显示
    /// </summary>
    wdPaneNone,

    /// <summary>
    /// 主要页眉窗格
    /// </summary>
    wdPanePrimaryHeader,

    /// <summary>
    /// 首页页眉
    /// </summary>
    wdPaneFirstPageHeader,

    /// <summary>
    /// 偶数页页眉
    /// </summary>
    wdPaneEvenPagesHeader,

    /// <summary>
    /// 主要页脚窗格
    /// </summary>
    wdPanePrimaryFooter,

    /// <summary>
    /// 首页页脚
    /// </summary>
    wdPaneFirstPageFooter,

    /// <summary>
    /// 偶数页页脚
    /// </summary>
    wdPaneEvenPagesFooter,

    /// <summary>
    /// 脚注
    /// </summary>
    wdPaneFootnotes,

    /// <summary>
    /// 尾注
    /// </summary>
    wdPaneEndnotes,

    /// <summary>
    /// 脚注接续通知
    /// </summary>
    wdPaneFootnoteContinuationNotice,

    /// <summary>
    /// 脚注接续分隔符
    /// </summary>
    wdPaneFootnoteContinuationSeparator,

    /// <summary>
    /// 脚注分隔符
    /// </summary>
    wdPaneFootnoteSeparator,

    /// <summary>
    /// 尾注接续通知
    /// </summary>
    wdPaneEndnoteContinuationNotice,

    /// <summary>
    /// 尾注接续分隔符
    /// </summary>
    wdPaneEndnoteContinuationSeparator,

    /// <summary>
    /// 尾注分隔符
    /// </summary>
    wdPaneEndnoteSeparator,

    /// <summary>
    /// 选定的批注
    /// </summary>
    wdPaneComments,

    /// <summary>
    /// 当前页页眉
    /// </summary>
    wdPaneCurrentPageHeader,

    /// <summary>
    /// 当前页页脚
    /// </summary>
    wdPaneCurrentPageFooter,

    /// <summary>
    /// 修订窗格
    /// </summary>
    wdPaneRevisions,

    /// <summary>
    /// 修订窗格显示在文档窗口底部
    /// </summary>
    wdPaneRevisionsHoriz,

    /// <summary>
    /// 修订窗格显示在文档窗口左侧
    /// </summary>
    wdPaneRevisionsVert
}