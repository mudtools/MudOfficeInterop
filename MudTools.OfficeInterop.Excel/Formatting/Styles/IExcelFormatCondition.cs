//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel FormatCondition 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.FormatCondition (及 ColorScale, DataBar, IconSetCondition) 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelFormatCondition : IOfficeObject<IExcelFormatCondition, MsExcel.FormatCondition>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取条件格式规则的父对象 (通常是 FormatConditions 集合)
    /// 对应 FormatCondition.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取条件格式规则所在的Application对象
    /// 对应 FormatCondition.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取条件格式规则的类型
    /// 对应 FormatCondition.Type 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlFormatConditionType Type { get; }

    /// <summary>
    /// 获取或设置比较操作符 (对于 xlCellValue 类型)
    /// 对应 FormatCondition.Operator 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlFormatConditionOperator Operator { get; }

    /// <summary>
    /// 获取或设置公式1
    /// 对应 FormatCondition.Formula1 属性
    /// </summary>
    string Formula1 { get; }

    /// <summary>
    /// 获取或设置公式2
    /// 对应 FormatCondition.Formula2 属性
    /// </summary>
    string Formula2 { get; }

    /// <summary>
    /// 获取或设置文本 (对于 xlTextString 类型)
    /// 对应 FormatCondition.Text 属性
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置文本比较器 (对于 xlTextString 类型)
    /// 对应 FormatCondition.TextOperator 属性
    /// </summary>
    XlContainsOperator TextOperator { get; set; }
    #endregion

    #region 格式设置 
    /// <summary>
    /// 获取条件格式规则的字体对象
    /// 对应 FormatCondition.Font 属性
    /// </summary>
    IExcelFont? Font { get; }

    /// <summary>
    /// 获取条件格式规则的背景对象
    /// 对应 FormatCondition.Interior 属性
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取条件格式规则的边框对象
    /// 对应 FormatCondition.Borders 属性
    /// </summary>
    IExcelBorders? Borders { get; }

    /// <summary>
    /// 获取或设置条件格式应用的单元格区域
    /// 对应 FormatCondition.AppliesTo 属性
    /// </summary>
    IExcelRange? AppliesTo { get; }

    /// <summary>
    /// 获取或设置条件格式规则的优先级
    /// 对应 FormatCondition.Priority 属性
    /// </summary>
    int Priority { get; set; }

    /// <summary>
    /// 获取或设置数据透视表条件格式的作用范围类型
    /// 对应 FormatCondition.ScopeType 属性
    /// </summary>
    XlPivotConditionScope ScopeType { get; set; }

    /// <summary>
    /// 获取或设置日期条件格式的时间段操作符
    /// 对应 FormatCondition.DateOperator 属性
    /// </summary>
    XlTimePeriods DateOperator { get; set; }

    /// <summary>
    /// 获取一个值，指示条件格式是否与数据透视表相关
    /// 对应 FormatCondition.PTCondition 属性
    /// </summary>
    bool PTCondition { get; }

    /// <summary>
    /// 获取或设置当条件为真时是否停止评估其他条件格式规则
    /// 对应 FormatCondition.StopIfTrue 属性
    /// </summary>
    bool StopIfTrue { get; set; }

    /// <summary>
    /// 获取或设置条件格式规则的编号格式
    /// 对应 FormatCondition.NumberFormat 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string? NumberFormat { get; set; }
    #endregion

    #region 操作方法

    /// <summary>
    /// 修改应用条件格式的单元格区域
    /// </summary>
    /// <param name="Range">要应用条件格式的新区域</param>
    void ModifyAppliesToRange(IExcelRange Range);

    /// <summary>
    /// 将条件格式规则设置为最高优先级
    /// </summary>
    void SetFirstPriority();

    /// <summary>
    /// 将条件格式规则设置为最低优先级
    /// </summary>
    void SetLastPriority();

    /// <summary>
    /// 删除此条件格式规则
    /// 对应 FormatCondition.Delete 方法
    /// </summary>
    void Delete();

    /// <summary>
    /// 修改此条件格式规则 (适用于 xlCellValue, xlExpression)
    /// 对应 FormatCondition.Modify 方法
    /// </summary>
    /// <param name="type">条件类型</param>
    /// <param name="cOperator">比较操作符</param>
    /// <param name="formula1">公式1</param>
    /// <param name="formula2">公式2</param>
    void Modify(XlFormatConditionType type, XlFormatConditionOperator cOperator, string formula1, string formula2);

    /// <summary>
    /// 修改此条件格式规则 (适用于 xlCellValue, xlExpression)
    /// 对应 FormatCondition.ModifyEx 方法
    /// </summary>
    void ModifyEx(XlFormatConditionType type, XlFormatConditionOperator cOperator1, object formula1, string formula2, object String, XlFormatConditionOperator cOperator2);

    #endregion

}
