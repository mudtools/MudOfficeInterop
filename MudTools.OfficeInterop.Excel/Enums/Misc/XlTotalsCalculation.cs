//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定在表格或列表对象的汇总行中显示的计算类型
/// </summary>
public enum XlTotalsCalculation
{
    /// <summary>
    /// 无计算
    /// </summary>
    xlTotalsCalculationNone,
    /// <summary>
    /// 求和计算
    /// </summary>
    xlTotalsCalculationSum,
    /// <summary>
    /// 平均值计算
    /// </summary>
    xlTotalsCalculationAverage,
    /// <summary>
    /// 计数计算
    /// </summary>
    xlTotalsCalculationCount,
    /// <summary>
    /// 数值计数计算
    /// </summary>
    xlTotalsCalculationCountNums,
    /// <summary>
    /// 最小值计算
    /// </summary>
    xlTotalsCalculationMin,
    /// <summary>
    /// 最大值计算
    /// </summary>
    xlTotalsCalculationMax,
    /// <summary>
    /// 标准偏差计算
    /// </summary>
    xlTotalsCalculationStdDev,
    /// <summary>
    /// 方差计算
    /// </summary>
    xlTotalsCalculationVar,
    /// <summary>
    /// 自定义计算
    /// </summary>
    xlTotalsCalculationCustom
}