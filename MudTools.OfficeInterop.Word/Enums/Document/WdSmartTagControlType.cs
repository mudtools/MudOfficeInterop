//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定智能标记控件的类型
/// </summary>
public enum WdSmartTagControlType
{
    /// <summary>
/// 智能标记控件
/// </summary>
    wdControlSmartTag = 1,
    /// <summary>
/// 链接控件
/// </summary>
    wdControlLink,
    /// <summary>
/// 帮助控件
/// </summary>
    wdControlHelp,
    /// <summary>
/// 帮助URL控件
/// </summary>
    wdControlHelpURL,
    /// <summary>
/// 分隔符控件
/// </summary>
    wdControlSeparator,
    /// <summary>
/// 按钮控件
/// </summary>
    wdControlButton,
    /// <summary>
/// 标签控件
/// </summary>
    wdControlLabel,
    /// <summary>
/// 图像控件
/// </summary>
    wdControlImage,
    /// <summary>
/// 复选框控件
/// </summary>
    wdControlCheckbox,
    /// <summary>
/// 文本框控件
/// </summary>
    wdControlTextbox,
    /// <summary>
/// 列表框控件
/// </summary>
    wdControlListbox,
    /// <summary>
/// 组合框控件
/// </summary>
    wdControlCombo,
    /// <summary>
/// ActiveX控件
/// </summary>
    wdControlActiveX,
    /// <summary>
/// 文档片段控件
/// </summary>
    wdControlDocumentFragment,
    /// <summary>
/// 文档片段URL控件
/// </summary>
    wdControlDocumentFragmentURL,
    /// <summary>
/// 单选按钮组控件
/// </summary>
    wdControlRadioGroup
}