//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定图表元素类型的枚举，用于标识和操作图表中的不同元素
/// </summary>
public enum MsoChartElementType
{
    /// <summary>
    /// 图表底座
    /// </summary>
    msoElementChartFloor = 1101,

    /// <summary>
    /// 图表背景墙
    /// </summary>
    msoElementChartWall = 1102,

    /// <summary>
    /// 系列轴
    /// </summary>
    msoElementSeriesAxis = 1103,

    /// <summary>
    /// 数值轴
    /// </summary>
    msoElementValueAxis = 1104,

    /// <summary>
    /// 分类轴
    /// </summary>
    msoElementCategoryAxis = 1105,

    /// <summary>
    /// 数据表
    /// </summary>
    msoElementDataTable = 1106,

    /// <summary>
    /// 图例
    /// </summary>
    msoElementLegend = 1107,

    /// <summary>
    /// 绘图区
    /// </summary>
    msoElementPlotArea = 1108,

    /// <summary>
    /// 主要分类轴标题
    /// </summary>
    msoElementPrimaryCategoryAxisTitle = 1109,

    /// <summary>
    /// 主要数值轴标题
    /// </summary>
    msoElementPrimaryValueAxisTitle = 1110,

    /// <summary>
    /// 主要分类轴
    /// </summary>
    msoElementPrimaryCategoryAxis = 1111,

    /// <summary>
    /// 主要数值轴
    /// </summary>
    msoElementPrimaryValueAxis = 1112,

    /// <summary>
    /// 次要分类轴
    /// </summary>
    msoElementSecondaryCategoryAxis = 1113,

    /// <summary>
    /// 次要数值轴
    /// </summary>
    msoElementSecondaryValueAxis = 1114,

    /// <summary>
    /// 次要分类轴标题
    /// </summary>
    msoElementSecondaryCategoryAxisTitle = 1115,

    /// <summary>
    /// 次要数值轴标题
    /// </summary>
    msoElementSecondaryValueAxisTitle = 1116,

    /// <summary>
    /// 系列轴标题
    /// </summary>
    msoElementSeriesAxisTitle = 1117,

    /// <summary>
    /// 主要分类网格线
    /// </summary>
    msoElementPrimaryCategoryGridLines = 1118,

    /// <summary>
    /// 主要数值网格线
    /// </summary>
    msoElementPrimaryValueGridLines = 1119,

    /// <summary>
    /// 次要分类网格线
    /// </summary>
    msoElementSecondaryCategoryGridLines = 1120,

    /// <summary>
    /// 次要数值网格线
    /// </summary>
    msoElementSecondaryValueGridLines = 1121,

    /// <summary>
    /// 系列轴网格线
    /// </summary>
    msoElementSeriesAxisGridLines = 1122,

    /// <summary>
    /// 无标签的主要分类轴
    /// </summary>
    msoElementPrimaryCategoryAxisWithoutLabels = 1123,

    /// <summary>
    /// 反向的主要分类轴
    /// </summary>
    msoElementPrimaryCategoryAxisReverse = 1124,

    /// <summary>
    /// 反向的主要数值轴
    /// </summary>
    msoElementPrimaryValueAxisReverse = 1125,

    /// <summary>
    /// 百万级的主要分类轴
    /// </summary>
    msoElementPrimaryCategoryAxisMillions = 1126,

    /// <summary>
    /// 百万级的主要数值轴
    /// </summary>
    msoElementPrimaryValueAxisMillions = 1127,

    /// <summary>
    /// 十亿级的主要分类轴
    /// </summary>
    msoElementPrimaryCategoryAxisBillions = 1128,

    /// <summary>
    /// 十亿级的主要数值轴
    /// </summary>
    msoElementPrimaryValueAxisBillions = 1129,

    /// <summary>
    /// 对数刻度的主要分类轴
    /// </summary>
    msoElementPrimaryCategoryAxisLogScale = 1130,

    /// <summary>
    /// 对数刻度的主要数值轴
    /// </summary>
    msoElementPrimaryValueAxisLogScale = 1131,

    /// <summary>
    /// 千级的主要分类轴
    /// </summary>
    msoElementPrimaryCategoryAxisThousands = 1132,

    /// <summary>
    /// 千级的主要数值轴
    /// </summary>
    msoElementPrimaryValueAxisThousands = 1133,

    /// <summary>
    /// 百级的主要分类轴
    /// </summary>
    msoElementPrimaryCategoryAxisHundreds = 1134,

    /// <summary>
    /// 百级的主要数值轴
    /// </summary>
    msoElementPrimaryValueAxisHundreds = 1135,

    /// <summary>
    /// 十级的主要分类轴
    /// </summary>
    msoElementPrimaryCategoryAxisTens = 1136,

    /// <summary>
    /// 十级的主要数值轴
    /// </summary>
    msoElementPrimaryValueAxisTens = 1137,

    /// <summary>
    /// 个位级的主要分类轴
    /// </summary>
    msoElementPrimaryCategoryAxisOnes = 1138,

    /// <summary>
    /// 个位级的主要数值轴
    /// </summary>
    msoElementPrimaryValueAxisOnes = 1139,

    /// <summary>
    /// 小数级的主要分类轴
    /// </summary>
    msoElementPrimaryCategoryAxisDecimal = 1140,

    /// <summary>
    /// 小数级的主要数值轴
    /// </summary>
    msoElementPrimaryValueAxisDecimal = 1141,

    /// <summary>
    /// 货币级的主要分类轴
    /// </summary>
    msoElementPrimaryCategoryAxisCurrency = 1142,

    /// <summary>
    /// 货币级的主要数值轴
    /// </summary>
    msoElementPrimaryValueAxisCurrency = 1143,

    /// <summary>
    /// 百分级的主要分类轴
    /// </summary>
    msoElementPrimaryCategoryAxisPercentage = 1144,

    /// <summary>
    /// 百分级的主要数值轴
    /// </summary>
    msoElementPrimaryValueAxisPercentage = 1145,

    /// <summary>
    /// 科学计数法的主要分类轴
    /// </summary>
    msoElementPrimaryCategoryAxisScientific = 1146,

    /// <summary>
    /// 科学计数法的主要数值轴
    /// </summary>
    msoElementPrimaryValueAxisScientific = 1147,

    /// <summary>
    /// 日期型的主要分类轴
    /// </summary>
    msoElementPrimaryCategoryAxisDate = 1148,

    /// <summary>
    /// 日期型的主要数值轴
    /// </summary>
    msoElementPrimaryValueAxisDate = 1149,

    /// <summary>
    /// 时间型的主要分类轴
    /// </summary>
    msoElementPrimaryCategoryAxisTime = 1150,

    /// <summary>
    /// 时间型的主要数值轴
    /// </summary>
    msoElementPrimaryValueAxisTime = 1151,

    /// <summary>
    /// 分数型的主要分类轴
    /// </summary>
    msoElementPrimaryCategoryAxisFraction = 1152,

    /// <summary>
    /// 分数型的主要数值轴
    /// </summary>
    msoElementPrimaryValueAxisFraction = 1153,

    /// <summary>
    /// 自定义的主要分类轴
    /// </summary>
    msoElementPrimaryCategoryAxisCustom = 1154,

    /// <summary>
    /// 自定义的主要数值轴
    /// </summary>
    msoElementPrimaryValueAxisCustom = 1155
}