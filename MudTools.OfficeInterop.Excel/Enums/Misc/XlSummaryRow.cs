//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 汇总行位置枚举
/// 用于指定在创建分级显示或数据透视表时，汇总行相对于明细行的位置
/// </summary>
public enum XlSummaryRow
{
    /// <summary>
    /// 汇总行在上方
    /// 汇总行显示在明细行的上方
    /// </summary>
    xlSummaryAbove,

    /// <summary>
    /// 汇总行在下方
    /// 汇总行显示在明细行的下方
    /// </summary>
    xlSummaryBelow
}

/// <summary>
/// 汇总列位置枚举
/// 用于指定在创建分级显示或数据透视表时，汇总列相对于明细列的位置
/// </summary>
public enum XlSummaryColumn
{

    /// <summary>
    /// 汇总列在左侧
    /// 汇总列显示在明细列的左侧
    /// </summary>
    xlSummaryOnLeft = -4131,

    /// <summary>
    /// 汇总列在右侧
    /// 汇总列显示在明细列的右侧
    /// </summary>
    xlSummaryOnRight = -4152
}