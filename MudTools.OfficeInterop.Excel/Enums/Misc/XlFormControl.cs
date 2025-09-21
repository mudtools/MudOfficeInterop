//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定可以在工作表上使用的表单控件类型
/// </summary>
public enum XlFormControl
{
    /// <summary>
    /// 命令按钮控件
    /// </summary>
    xlButtonControl,

    /// <summary>
    /// 复选框控件
    /// </summary>
    xlCheckBox,

    /// <summary>
    /// 下拉列表控件
    /// </summary>
    xlDropDown,

    /// <summary>
    /// 文本输入框控件
    /// </summary>
    xlEditBox,

    /// <summary>
    /// 分组框控件
    /// </summary>
    xlGroupBox,

    /// <summary>
    /// 标签控件
    /// </summary>
    xlLabel,

    /// <summary>
    /// 列表框控件
    /// </summary>
    xlListBox,

    /// <summary>
    /// 选项按钮控件
    /// </summary>
    xlOptionButton,

    /// <summary>
    /// 滚动条控件
    /// </summary>
    xlScrollBar,

    /// <summary>
    /// 数值调节钮控件
    /// </summary>
    xlSpinner
}