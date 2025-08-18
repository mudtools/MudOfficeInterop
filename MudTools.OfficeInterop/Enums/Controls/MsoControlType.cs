//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定Office控件的类型
/// </summary>
public enum MsoControlType
{
    /// <summary>
    /// 自定义控件
    /// </summary>
    msoControlCustom,
    /// <summary>
    /// 命令按钮控件
    /// </summary>
    msoControlButton,
    /// <summary>
    /// 文本框控件
    /// </summary>
    msoControlEdit,
    /// <summary>
    /// 下拉列表控件
    /// </summary>
    msoControlDropdown,
    /// <summary>
    /// 组合框控件
    /// </summary>
    msoControlComboBox,
    /// <summary>
    /// 按钮下拉列表控件
    /// </summary>
    msoControlButtonDropdown,
    /// <summary>
    /// 分割下拉列表控件
    /// </summary>
    msoControlSplitDropdown,
    /// <summary>
    /// OCX下拉列表控件
    /// </summary>
    msoControlOCXDropdown,
    /// <summary>
    /// 通用下拉列表控件
    /// </summary>
    msoControlGenericDropdown,
    /// <summary>
    /// 图形下拉列表控件
    /// </summary>
    msoControlGraphicDropdown,
    /// <summary>
    /// 弹出菜单控件
    /// </summary>
    msoControlPopup,
    /// <summary>
    /// 图形弹出菜单控件
    /// </summary>
    msoControlGraphicPopup,
    /// <summary>
    /// 按钮弹出菜单控件
    /// </summary>
    msoControlButtonPopup,
    /// <summary>
    /// 分割按钮弹出菜单控件
    /// </summary>
    msoControlSplitButtonPopup,
    /// <summary>
    /// 分割按钮最近使用弹出菜单控件
    /// </summary>
    msoControlSplitButtonMRUPopup,
    /// <summary>
    /// 标签控件
    /// </summary>
    msoControlLabel,
    /// <summary>
    /// 可扩展网格控件
    /// </summary>
    msoControlExpandingGrid,
    /// <summary>
    /// 分割可扩展网格控件
    /// </summary>
    msoControlSplitExpandingGrid,
    /// <summary>
    /// 网格控件
    /// </summary>
    msoControlGrid,
    /// <summary>
    /// 仪表控件
    /// </summary>
    msoControlGauge,
    /// <summary>
    /// 图形组合框控件
    /// </summary>
    msoControlGraphicCombo,
    /// <summary>
    /// 窗格控件
    /// </summary>
    msoControlPane,
    /// <summary>
    /// ActiveX控件
    /// </summary>
    msoControlActiveX,
    /// <summary>
    /// 微调控件
    /// </summary>
    msoControlSpinner,
    /// <summary>
    /// 扩展标签控件
    /// </summary>
    msoControlLabelEx,
    /// <summary>
    /// 工作窗格控件
    /// </summary>
    msoControlWorkPane,
    /// <summary>
    /// 自动完成组合框控件
    /// </summary>
    msoControlAutoCompleteCombo
}