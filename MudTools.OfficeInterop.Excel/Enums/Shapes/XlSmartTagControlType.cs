//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定智能标签控件的类型
/// </summary>
public enum XlSmartTagControlType
{
    /// <summary>
    /// 智能标签控件
    /// </summary>
    xlSmartTagControlSmartTag = 1,

    /// <summary>
    /// 链接控件
    /// </summary>
    xlSmartTagControlLink,

    /// <summary>
    /// 帮助控件
    /// </summary>
    xlSmartTagControlHelp,

    /// <summary>
    /// 帮助URL控件
    /// </summary>
    xlSmartTagControlHelpURL,

    /// <summary>
    /// 分隔符控件
    /// </summary>
    xlSmartTagControlSeparator,

    /// <summary>
    /// 按钮控件
    /// </summary>
    xlSmartTagControlButton,

    /// <summary>
    /// 标签控件
    /// </summary>
    xlSmartTagControlLabel,

    /// <summary>
    /// 图像控件
    /// </summary>
    xlSmartTagControlImage,

    /// <summary>
    /// 复选框控件
    /// </summary>
    xlSmartTagControlCheckbox,

    /// <summary>
    /// 文本框控件
    /// </summary>
    xlSmartTagControlTextbox,

    /// <summary>
    /// 列表框控件
    /// </summary>
    xlSmartTagControlListbox,

    /// <summary>
    /// 组合框控件
    /// </summary>
    xlSmartTagControlCombo,

    /// <summary>
    /// ActiveX控件
    /// </summary>
    xlSmartTagControlActiveX,

    /// <summary>
    /// 单选按钮组控件
    /// </summary>
    xlSmartTagControlRadioGroup
}