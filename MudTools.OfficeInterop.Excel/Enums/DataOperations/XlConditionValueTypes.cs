//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 条件值类型枚举
/// 用于指定条件格式中条件值的类型
/// </summary>
public enum XlConditionValueTypes
{
    /// <summary>
    /// 无值
    /// 不指定条件值
    /// </summary>
    xlConditionValueNone = -1,
    
    /// <summary>
    /// 数字值
    /// 条件值为具体的数字
    /// </summary>
    xlConditionValueNumber,
    
    /// <summary>
    /// 最低值
    /// 条件值为数据范围中的最低值
    /// </summary>
    xlConditionValueLowestValue,
    
    /// <summary>
    /// 最高值
    /// 条件值为数据范围中的最高值
    /// </summary>
    xlConditionValueHighestValue,
    
    /// <summary>
    /// 百分比
    /// 条件值为数据范围中的百分比
    /// </summary>
    xlConditionValuePercent,
    
    /// <summary>
    /// 公式
    /// 条件值由公式计算得出
    /// </summary>
    xlConditionValueFormula,
    
    /// <summary>
    /// 百分点
    /// 条件值为数据范围中的百分点值
    /// </summary>
    xlConditionValuePercentile,
    
    /// <summary>
    /// 自动最小值
    /// 条件值为Excel自动确定的最小值
    /// </summary>
    xlConditionValueAutomaticMin,
    
    /// <summary>
    /// 自动最大值
    /// 条件值为Excel自动确定的最大值
    /// </summary>
    xlConditionValueAutomaticMax
}