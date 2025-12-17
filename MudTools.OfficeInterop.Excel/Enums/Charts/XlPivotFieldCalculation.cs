//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// 指定使用自定义计算时数据透视表字段执行的计算类型
/// </summary>
public enum XlPivotFieldCalculation
{
    /// <summary>
    /// 与基字段中基项值的差值
    /// </summary>
    xlDifferenceFrom = 2,

    /// <summary>
    /// 数据计算为 ((单元格中的值) x (总计的总计)) / ((行总计) x (列总计))
    /// </summary>
    xlIndex = 9,

    /// <summary>
    /// 无额外计算
    /// </summary>
    xlNoAdditionalCalculation = -4143,

    /// <summary>
    /// 与基字段中基项值的百分比差值
    /// </summary>
    xlPercentDifferenceFrom = 4,

    /// <summary>
    /// 占基字段中基项值的百分比
    /// </summary>
    xlPercentOf = 3,

    /// <summary>
    /// 占列或系列总计的百分比
    /// </summary>
    xlPercentOfColumn = 7,

    /// <summary>
    /// 占行或分类总计的百分比
    /// </summary>
    xlPercentOfRow = 6,

    /// <summary>
    /// 占报表中所有数据或数据点总和的百分比
    /// </summary>
    xlPercentOfTotal = 8,

    /// <summary>
    /// 将基字段中连续项的数据作为累计总计
    /// </summary>
    xlRunningTotal = 5,

    /// <summary>
    /// 占父行总计的百分比
    /// </summary>
    xlPercentOfParentRow = 10,

    /// <summary>
    /// 占父列总计的百分比
    /// </summary>
    xlPercentOfParentColumn = 11,

    /// <summary>
    /// 占指定父基字段总计的百分比
    /// </summary>
    xlPercentOfParent = 12,

    /// <summary>
    /// 占指定基字段累计总计的百分比
    /// </summary>
    xlPercentRunningTotal = 13,

    /// <summary>
    /// 按从小到大排序
    /// </summary>
    xlRankAscending = 14,

    /// <summary>
    /// 按从大到小排序
    /// </summary>
    xlRankDecending = 15
}