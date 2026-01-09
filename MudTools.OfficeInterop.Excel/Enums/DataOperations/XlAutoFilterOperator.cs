//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定自动筛选操作符的枚举类型
/// </summary>
public enum XlAutoFilterOperator
{
    /// <summary>
    /// 逻辑与操作符，用于连接两个条件
    /// </summary>
    xlAnd = 1,

    /// <summary>
    /// 筛选底部10项
    /// </summary>
    xlBottom10Items = 4,

    /// <summary>
    /// 筛选底部10百分比
    /// </summary>
    xlBottom10Percent = 6,

    /// <summary>
    /// 逻辑或操作符，用于连接两个条件中的任一个
    /// </summary>
    xlOr = 2,

    /// <summary>
    /// 筛选顶部10项
    /// </summary>
    xlTop10Items = 3,

    /// <summary>
    /// 筛选顶部10百分比
    /// </summary>
    xlTop10Percent = 5,

    /// <summary>
    /// 按值筛选
    /// </summary>
    xlFilterValues = 7,

    /// <summary>
    /// 按单元格颜色筛选
    /// </summary>
    xlFilterCellColor = 8,

    /// <summary>
    /// 按字体颜色筛选
    /// </summary>
    xlFilterFontColor = 9,

    /// <summary>
    /// 按图标筛选
    /// </summary>
    xlFilterIcon = 10,

    /// <summary>
    /// 动态筛选
    /// </summary>
    xlFilterDynamic = 11,

    /// <summary>
    /// 筛选无填充内容
    /// </summary>
    xlFilterNoFill = 12,

    /// <summary>
    /// 筛选自动字体颜色
    /// </summary>
    xlFilterAutomaticFontColor = 13,

    /// <summary>
    /// 筛选无图标
    /// </summary>
    xlFilterNoIcon = 14
}