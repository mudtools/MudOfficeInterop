//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel Databar 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Databar 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelDatabar : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取数据条对象的父对象 (通常是 FormatCondition)
    /// 对应 Databar.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取数据条对象所在的Application对象
    /// 对应 Databar.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取应用数据条格式的单元格范围
    /// 对应 Databar.AppliesTo 属性
    /// </summary>
    IExcelRange? AppliesTo { get; }

    /// <summary>
    /// 获取数据条的颜色设置
    /// 对应 Databar.BarColor 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelFormatColor? BarColor { get; }

    /// <summary>
    /// 获取或设置最小条件值
    /// 对应 Databar.MinPoint 属性
    /// </summary>
    IExcelConditionValue? MinPoint { get; }

    /// <summary>
    /// 获取或设置最大条件值
    /// 对应 Databar.MaxPoint 属性
    /// </summary>
    IExcelConditionValue? MaxPoint { get; }

    /// <summary>
    /// 获取或设置数据条的方向
    /// 对应 Databar.Direction 属性
    /// </summary>
    int Direction { get; set; }

    /// <summary>
    /// 获取或设置数据条的图形条显示
    /// 对应 Databar.BarFillType 属性
    /// </summary>
    XlDataBarFillType BarFillType { get; set; }

    /// <summary>
    /// 获取或设置数据透视表条件格式的作用范围
    /// 对应 Databar.ScopeType 属性
    /// </summary>
    XlPivotConditionScope ScopeType { get; set; }

    /// <summary>
    /// 获取或设置数据条的轴位置
    /// 对应 Databar.AxisPosition 属性
    /// </summary>
    XlDataBarAxisPosition AxisPosition { get; set; }

    /// <summary>
    /// 获取数据条轴的颜色
    /// 对应 Databar.AxisColor 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color? AxisColor { get; }

    /// <summary>
    /// 获取一个值，指示条件格式是否与数据透视表相关
    /// 对应 Databar.PTCondition 属性
    /// </summary>
    bool PTCondition { get; }

    /// <summary>
    /// 获取或设置数据条的最大百分比值
    /// 对应 Databar.PercentMax 属性
    /// </summary>
    int PercentMax { get; set; }

    /// <summary>
    /// 获取或设置数据条的最小百分比值
    /// 对应 Databar.PercentMin 属性
    /// </summary>
    int PercentMin { get; set; }

    /// <summary>
    /// 获取或设置条件格式规则的优先级
    /// 对应 Databar.Priority 属性
    /// </summary>
    int Priority { get; set; }

    /// <summary>
    /// 获取或设置当条件为真时是否停止评估其他条件格式规则
    /// 对应 Databar.StopIfTrue 属性
    /// </summary>
    bool StopIfTrue { get; }
    #endregion

    #region 格式设置


    IExcelDataBarBorder? BarBorder { get; }

    int Type { get; }

    bool ShowValue { get; set; }

    /// <summary>
    /// 获取负值数据条的格式设置对象
    /// </summary>
    IExcelNegativeBarFormat? NegativeBarFormat { get; }

    /// <summary>
    /// 获取数据条的字体对象
    /// </summary>
    string Formula { get; set; }
    #endregion

    /// <summary>
    /// 修改应用数据条格式的单元格范围
    /// </summary>
    /// <param name="Range">新的单元格范围</param>
    void ModifyAppliesToRange(IExcelRange Range);

    /// <summary>
    /// 删除数据条条件格式规则
    /// </summary>
    void Delete();

    /// <summary>
    /// 将数据条条件格式规则的优先级设置为最低
    /// </summary>
    void SetLastPriority();

    /// <summary>
    /// 将数据条条件格式规则的优先级设置为最高
    /// </summary>
    void SetFirstPriority();
}
