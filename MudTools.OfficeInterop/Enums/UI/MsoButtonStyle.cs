//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定命令栏按钮的显示样式
/// </summary>
public enum MsoButtonStyle
{
    /// <summary>
    /// 自动样式 - 系统决定按钮显示方式
    /// </summary>
    msoButtonAutomatic = 0,

    /// <summary>
    /// 仅显示图标
    /// </summary>
    msoButtonIcon = 1,

    /// <summary>
    /// 仅显示文字标题
    /// </summary>
    msoButtonCaption = 2,

    /// <summary>
    /// 显示图标和文字标题（文字在图标右侧）
    /// </summary>
    msoButtonIconAndCaption = 3,

    /// <summary>
    /// 显示图标和自动换行的文字标题（文字在图标右侧）
    /// </summary>
    msoButtonIconAndWrapCaption = 7,

    /// <summary>
    /// 显示图标和文字标题（文字在图标下方）
    /// </summary>
    msoButtonIconAndCaptionBelow = 11,

    /// <summary>
    /// 仅显示自动换行的文字标题
    /// </summary>
    msoButtonWrapCaption = 14,

    /// <summary>
    /// 显示图标和自动换行的文字标题（文字在图标下方）
    /// </summary>
    msoButtonIconAndWrapCaptionBelow = 15
}