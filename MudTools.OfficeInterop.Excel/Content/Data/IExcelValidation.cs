//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel Validation 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Validation 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelValidation : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取数据验证规则的父对象 (通常是 Range)
    /// 对应 Validation.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取数据验证规则所在的Application对象
    /// 对应 Validation.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置验证类型
    /// 对应 Validation.Type 属性
    /// </summary>
    XlDVType Type { get; }

    /// <summary>
    /// 获取或设置错误警告样式
    /// 对应 Validation.AlertStyle 属性
    /// </summary>
    XlDVAlertStyle AlertStyle { get; }

    /// <summary>
    /// 获取或设置公式1
    /// 对应 Validation.Formula1 属性
    /// </summary>
    string Formula1 { get; }

    /// <summary>
    /// 获取或设置公式2
    /// 对应 Validation.Formula2 属性
    /// </summary>
    string Formula2 { get; }

    /// <summary>
    /// 获取或设置是否为值
    /// 对应 Validation.Value 属性 (或 Add 方法的参数)
    /// </summary>
    bool Value { get; }

    /// <summary>
    /// 获取或设置输入提示标题
    /// 对应 Validation.InputTitle 属性
    /// </summary>
    string InputTitle { get; set; }

    /// <summary>
    /// 获取或设置输入提示信息
    /// 对应 Validation.InputMessage 属性
    /// </summary>
    string InputMessage { get; set; }

    /// <summary>
    /// 获取或设置是否显示错误提示
    /// 对应 Validation.ShowError 属性
    /// </summary>
    bool ShowError { get; set; }

    /// <summary>
    /// 获取或设置是否显示输入提示
    /// 对应 Validation.ShowInput 属性
    /// </summary>
    bool ShowInput { get; set; }

    /// <summary>
    /// 获取或设置错误提示标题
    /// 对应 Validation.ErrorTitle 属性
    /// </summary>
    string ErrorTitle { get; set; }

    /// <summary>
    /// 获取或设置错误提示信息
    /// 对应 Validation.ErrorMessage 属性
    /// </summary>
    string ErrorMessage { get; set; }

    /// <summary>
    /// 获取或设置是否忽略空值
    /// 对应 Validation.IgnoreBlank 属性
    /// </summary>
    bool IgnoreBlank { get; set; }

    /// <summary>
    /// 获取或设置是否在单元格内显示下拉箭头
    /// 对应 Validation.InCellDropdown 属性
    /// </summary>
    bool InCellDropdown { get; set; }
    #endregion

    #region 操作方法
    /// <summary>
    /// 删除此数据验证规则
    /// 对应 Validation.Delete 方法
    /// </summary>
    void Delete();

    /// <summary>
    /// 添加数据验证规则
    /// </summary>
    /// <param name="type">验证类型</param>
    /// <param name="alertStyle">警告样式</param>
    /// <param name="conditionOperator">数据验证运算符。</param>
    /// <param name="formula1"> 数据验证公式中的第一部分。</param>
    /// <param name="formula2">当 为 xlBetween 或 xlNotBetween 时Operator，数据验证的第二部分 (否则，将忽略此参数) 。</param>
    void Add(XlDVType type,
        XlDVAlertStyle? alertStyle = null,
        XlFormatConditionOperator? conditionOperator = null,
        string? formula1 = null,
        string? formula2 = null);

    /// <summary>
    /// 修改指定区域的数据有效性验证。
    /// </summary>
    /// <param name="type">验证类型</param>
    /// <param name="alertStyle">警告样式</param>
    /// <param name="formula1">数据验证公式中的第一部分。</param>
    /// <param name="formula2">当 为 xlBetween 或 xlNotBetween 时Operator，数据验证的第二部分 (否则，将忽略此参数) 。</param>
    /// <param name="formatConditionOperator">数据验证运算符。</param>
    void Modify(
        XlDVType type,
        XlDVAlertStyle? alertStyle = null,
        XlFormatConditionOperator? formatConditionOperator = null,
        string? formula1 = null, string? formula2 = null);
    #endregion
}