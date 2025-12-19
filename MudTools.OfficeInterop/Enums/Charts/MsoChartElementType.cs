//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 图表元素类型枚举，用于指定图表中各种元素的显示方式和位置
/// </summary>
public enum MsoChartElementType
{
    // 图表标题设置
    /// <summary>
    /// 不显示图表标题
    /// </summary>
    msoElementChartTitleNone = 0,
    /// <summary>
    /// 图表标题居中覆盖在图表上方
    /// </summary>
    msoElementChartTitleCenteredOverlay = 1,
    /// <summary>
    /// 图表标题显示在图表上方
    /// </summary>
    msoElementChartTitleAboveChart = 2,

    // 图例设置 (100-106)
    /// <summary>
    /// 不显示图例
    /// </summary>
    msoElementLegendNone = 100,
    /// <summary>
    /// 图例显示在右侧
    /// </summary>
    msoElementLegendRight = 101,
    /// <summary>
    /// 图例显示在顶部
    /// </summary>
    msoElementLegendTop = 102,
    /// <summary>
    /// 图例显示在左侧
    /// </summary>
    msoElementLegendLeft = 103,
    /// <summary>
    /// 图例显示在底部
    /// </summary>
    msoElementLegendBottom = 104,
    /// <summary>
    /// 图例以覆盖方式显示在右侧
    /// </summary>
    msoElementLegendRightOverlay = 105,
    /// <summary>
    /// 图例以覆盖方式显示在左侧
    /// </summary>
    msoElementLegendLeftOverlay = 106,

    // 数据标签设置 (200-211)
    /// <summary>
    /// 不显示数据标签
    /// </summary>
    msoElementDataLabelNone = 200,
    /// <summary>
    /// 显示数据标签
    /// </summary>
    msoElementDataLabelShow = 201,
    /// <summary>
    /// 数据标签居中显示
    /// </summary>
    msoElementDataLabelCenter = 202,
    /// <summary>
    /// 数据标签显示在内部末端
    /// </summary>
    msoElementDataLabelInsideEnd = 203,
    /// <summary>
    /// 数据标签显示在内部基部
    /// </summary>
    msoElementDataLabelInsideBase = 204,
    /// <summary>
    /// 数据标签显示在外部末端
    /// </summary>
    msoElementDataLabelOutSideEnd = 205,
    /// <summary>
    /// 数据标签显示在左侧
    /// </summary>
    msoElementDataLabelLeft = 206,
    /// <summary>
    /// 数据标签显示在右侧
    /// </summary>
    msoElementDataLabelRight = 207,
    /// <summary>
    /// 数据标签显示在顶部
    /// </summary>
    msoElementDataLabelTop = 208,
    /// <summary>
    /// 数据标签显示在底部
    /// </summary>
    msoElementDataLabelBottom = 209,
    /// <summary>
    /// 数据标签最佳适应位置显示
    /// </summary>
    msoElementDataLabelBestFit = 210,
    /// <summary>
    /// 数据标签以标注线形式显示
    /// </summary>
    msoElementDataLabelCallout = 211,

    // 坐标轴标题设置 (300-327)
    /// <summary>
    /// 不显示主分类坐标轴标题
    /// </summary>
    msoElementPrimaryCategoryAxisTitleNone = 300,
    /// <summary>
    /// 主分类坐标轴标题紧邻坐标轴显示
    /// </summary>
    msoElementPrimaryCategoryAxisTitleAdjacentToAxis = 301,
    /// <summary>
    /// 主分类坐标轴标题显示在坐标轴下方
    /// </summary>
    msoElementPrimaryCategoryAxisTitleBelowAxis = 302,
    /// <summary>
    /// 主分类坐标轴标题旋转显示
    /// </summary>
    msoElementPrimaryCategoryAxisTitleRotated = 303,
    /// <summary>
    /// 主分类坐标轴标题垂直显示
    /// </summary>
    msoElementPrimaryCategoryAxisTitleVertical = 304,
    /// <summary>
    /// 主分类坐标轴标题水平显示
    /// </summary>
    msoElementPrimaryCategoryAxisTitleHorizontal = 305,
    /// <summary>
    /// 不显示主数值坐标轴标题
    /// </summary>
    msoElementPrimaryValueAxisTitleNone = 306,
    /// <summary>
    /// 主数值坐标轴标题紧邻坐标轴显示
    /// </summary>
    msoElementPrimaryValueAxisTitleAdjacentToAxis = 306,
    /// <summary>
    /// 主数值坐标轴标题显示在坐标轴下方
    /// </summary>
    msoElementPrimaryValueAxisTitleBelowAxis = 308,
    /// <summary>
    /// 主数值坐标轴标题旋转显示
    /// </summary>
    msoElementPrimaryValueAxisTitleRotated = 309,
    /// <summary>
    /// 主数值坐标轴标题垂直显示
    /// </summary>
    msoElementPrimaryValueAxisTitleVertical = 310,
    /// <summary>
    /// 主数值坐标轴标题水平显示
    /// </summary>
    msoElementPrimaryValueAxisTitleHorizontal = 311,
    /// <summary>
    /// 不显示次分类坐标轴标题
    /// </summary>
    msoElementSecondaryCategoryAxisTitleNone = 312,
    /// <summary>
    /// 次分类坐标轴标题紧邻坐标轴显示
    /// </summary>
    msoElementSecondaryCategoryAxisTitleAdjacentToAxis = 313,
    /// <summary>
    /// 次分类坐标轴标题显示在坐标轴下方
    /// </summary>
    msoElementSecondaryCategoryAxisTitleBelowAxis = 314,
    /// <summary>
    /// 次分类坐标轴标题旋转显示
    /// </summary>
    msoElementSecondaryCategoryAxisTitleRotated = 315,
    /// <summary>
    /// 次分类坐标轴标题垂直显示
    /// </summary>
    msoElementSecondaryCategoryAxisTitleVertical = 316,
    /// <summary>
    /// 次分类坐标轴标题水平显示
    /// </summary>
    msoElementSecondaryCategoryAxisTitleHorizontal = 317,
    /// <summary>
    /// 不显示次数值坐标轴标题
    /// </summary>
    msoElementSecondaryValueAxisTitleNone = 318,
    /// <summary>
    /// 次数值坐标轴标题紧邻坐标轴显示
    /// </summary>
    msoElementSecondaryValueAxisTitleAdjacentToAxis = 319,
    /// <summary>
    /// 次数值坐标轴标题显示在坐标轴下方
    /// </summary>
    msoElementSecondaryValueAxisTitleBelowAxis = 320,
    /// <summary>
    /// 次数值坐标轴标题旋转显示
    /// </summary>
    msoElementSecondaryValueAxisTitleRotated = 321,
    /// <summary>
    /// 次数值坐标轴标题垂直显示
    /// </summary>
    msoElementSecondaryValueAxisTitleVertical = 322,
    /// <summary>
    /// 次数值坐标轴标题水平显示
    /// </summary>
    msoElementSecondaryValueAxisTitleHorizontal = 323,
    /// <summary>
    /// 不显示系列坐标轴标题
    /// </summary>
    msoElementSeriesAxisTitleNone = 324,
    /// <summary>
    /// 系列坐标轴标题旋转显示
    /// </summary>
    msoElementSeriesAxisTitleRotated = 325,
    /// <summary>
    /// 系列坐标轴标题垂直显示
    /// </summary>
    msoElementSeriesAxisTitleVertical = 326,
    /// <summary>
    /// 系列坐标轴标题水平显示
    /// </summary>
    msoElementSeriesAxisTitleHorizontal = 327,

    // 网格线设置 (328-347)
    /// <summary>
    /// 不显示主数值网格线
    /// </summary>
    msoElementPrimaryValueGridLinesNone = 328,
    /// <summary>
    /// 显示主数值次要网格线
    /// </summary>
    msoElementPrimaryValueGridLinesMinor = 329,
    /// <summary>
    /// 显示主数值主要网格线
    /// </summary>
    msoElementPrimaryValueGridLinesMajor = 330,
    /// <summary>
    /// 显示主数值主要和次要网格线
    /// </summary>
    msoElementPrimaryValueGridLinesMinorMajor = 331,
    /// <summary>
    /// 不显示主分类网格线
    /// </summary>
    msoElementPrimaryCategoryGridLinesNone = 332,
    /// <summary>
    /// 显示主分类次要网格线
    /// </summary>
    msoElementPrimaryCategoryGridLinesMinor = 333,
    /// <summary>
    /// 显示主分类主要网格线
    /// </summary>
    msoElementPrimaryCategoryGridLinesMajor = 334,
    /// <summary>
    /// 显示主分类主要和次要网格线
    /// </summary>
    msoElementPrimaryCategoryGridLinesMinorMajor = 335,
    /// <summary>
    /// 不显示次数值网格线
    /// </summary>
    msoElementSecondaryValueGridLinesNone = 336,
    /// <summary>
    /// 显示次数值次要网格线
    /// </summary>
    msoElementSecondaryValueGridLinesMinor = 337,
    /// <summary>
    /// 显示次数值主要网格线
    /// </summary>
    msoElementSecondaryValueGridLinesMajor = 338,
    /// <summary>
    /// 显示次数值主要和次要网格线
    /// </summary>
    msoElementSecondaryValueGridLinesMinorMajor = 339,
    /// <summary>
    /// 不显示次分类网格线
    /// </summary>
    msoElementSecondaryCategoryGridLinesNone = 340,
    /// <summary>
    /// 显示次分类次要网格线
    /// </summary>
    msoElementSecondaryCategoryGridLinesMinor = 341,
    /// <summary>
    /// 显示次分类主要网格线
    /// </summary>
    msoElementSecondaryCategoryGridLinesMajor = 342,
    /// <summary>
    /// 显示次分类主要和次要网格线
    /// </summary>
    msoElementSecondaryCategoryGridLinesMinorMajor = 343,
    /// <summary>
    /// 不显示系列轴网格线
    /// </summary>
    msoElementSeriesAxisGridLinesNone = 344,
    /// <summary>
    /// 显示系列轴次要网格线
    /// </summary>
    msoElementSeriesAxisGridLinesMinor = 345,
    /// <summary>
    /// 显示系列轴主要网格线
    /// </summary>
    msoElementSeriesAxisGridLinesMajor = 346,
    /// <summary>
    /// 显示系列轴主要和次要网格线
    /// </summary>
    msoElementSeriesAxisGridLinesMinorMajor = 347,

    // 坐标轴设置 (348-379)
    /// <summary>
    /// 不显示主分类坐标轴
    /// </summary>
    msoElementPrimaryCategoryAxisNone = 348,
    /// <summary>
    /// 显示主分类坐标轴
    /// </summary>
    msoElementPrimaryCategoryAxisShow = 349,
    /// <summary>
    /// 显示无标签的主分类坐标轴
    /// </summary>
    msoElementPrimaryCategoryAxisWithoutLabels = 350,
    /// <summary>
    /// 反向显示主分类坐标轴
    /// </summary>
    msoElementPrimaryCategoryAxisReverse = 351,
    /// <summary>
    /// 不显示主数值坐标轴
    /// </summary>
    msoElementPrimaryValueAxisNone = 352,
    /// <summary>
    /// 显示主数值坐标轴
    /// </summary>
    msoElementPrimaryValueAxisShow = 353,
    /// <summary>
    /// 主数值坐标轴以千为单位显示
    /// </summary>
    msoElementPrimaryValueAxisThousands = 354,
    /// <summary>
    /// 主数值坐标轴以百万为单位显示
    /// </summary>
    msoElementPrimaryValueAxisMillions = 355,
    /// <summary>
    /// 主数值坐标轴以十亿为单位显示
    /// </summary>
    msoElementPrimaryValueAxisBillions = 356,
    /// <summary>
    /// 主数值坐标轴使用对数刻度
    /// </summary>
    msoElementPrimaryValueAxisLogScale = 357,
    /// <summary>
    /// 不显示次分类坐标轴
    /// </summary>
    msoElementSecondaryCategoryAxisNone = 358,
    /// <summary>
    /// 显示次分类坐标轴
    /// </summary>
    msoElementSecondaryCategoryAxisShow = 359,
    /// <summary>
    /// 显示无标签的次分类坐标轴
    /// </summary>
    msoElementSecondaryCategoryAxisWithoutLabels = 360,
    /// <summary>
    /// 反向显示次分类坐标轴
    /// </summary>
    msoElementSecondaryCategoryAxisReverse = 361,
    /// <summary>
    /// 不显示次数值坐标轴
    /// </summary>
    msoElementSecondaryValueAxisNone = 362,
    /// <summary>
    /// 显示次数值坐标轴
    /// </summary>
    msoElementSecondaryValueAxisShow = 363,
    /// <summary>
    /// 次数值坐标轴以千为单位显示
    /// </summary>
    msoElementSecondaryValueAxisThousands = 364,
    /// <summary>
    /// 次数值坐标轴以百万为单位显示
    /// </summary>
    msoElementSecondaryValueAxisMillions = 365,
    /// <summary>
    /// 次数值坐标轴以十亿为单位显示
    /// </summary>
    msoElementSecondaryValueAxisBillions = 366,
    /// <summary>
    /// 次数值坐标轴使用对数刻度
    /// </summary>
    msoElementSecondaryValueAxisLogScale = 367,
    /// <summary>
    /// 不显示系列坐标轴
    /// </summary>
    msoElementSeriesAxisNone = 368,
    /// <summary>
    /// 显示系列坐标轴
    /// </summary>
    msoElementSeriesAxisShow = 369,
    /// <summary>
    /// 显示无标签的系列坐标轴
    /// </summary>
    msoElementSeriesAxisWithoutLabeling = 370,
    /// <summary>
    /// 反向显示系列坐标轴
    /// </summary>
    msoElementSeriesAxisReverse = 371,
    /// <summary>
    /// 主分类坐标轴以千为单位显示
    /// </summary>
    msoElementPrimaryCategoryAxisThousands = 372,
    /// <summary>
    /// 主分类坐标轴以百万为单位显示
    /// </summary>
    msoElementPrimaryCategoryAxisMillions = 373,
    /// <summary>
    /// 主分类坐标轴以十亿为单位显示
    /// </summary>
    msoElementPrimaryCategoryAxisBillions = 374,
    /// <summary>
    /// 主分类坐标轴使用对数刻度
    /// </summary>
    msoElementPrimaryCategoryAxisLogScale = 375,
    /// <summary>
    /// 次分类坐标轴以千为单位显示
    /// </summary>
    msoElementSecondaryCategoryAxisThousands = 376,
    /// <summary>
    /// 次分类坐标轴以百万为单位显示
    /// </summary>
    msoElementSecondaryCategoryAxisMillions = 377,
    /// <summary>
    /// 次分类坐标轴以十亿为单位显示
    /// </summary>
    msoElementSecondaryCategoryAxisBillions = 378,
    /// <summary>
    /// 次分类坐标轴使用对数刻度
    /// </summary>
    msoElementSecondaryCategoryAxisLogScale = 379,

    // 数据表设置 (500-502)
    /// <summary>
    /// 不显示数据表
    /// </summary>
    msoElementDataTableNone = 500,
    /// <summary>
    /// 显示数据表
    /// </summary>
    msoElementDataTableShow = 501,
    /// <summary>
    /// 显示带图例键的数据表
    /// </summary>
    msoElementDataTableWithLegendKeys = 502,

    // 趋势线设置 (600-604)
    /// <summary>
    /// 不显示趋势线
    /// </summary>
    msoElementTrendlineNone = 600,
    /// <summary>
    /// 添加线性趋势线
    /// </summary>
    msoElementTrendlineAddLinear = 601,
    /// <summary>
    /// 添加指数趋势线
    /// </summary>
    msoElementTrendlineAddExponential = 602,
    /// <summary>
    /// 添加线性预测趋势线
    /// </summary>
    msoElementTrendlineAddLinearForecast = 603,
    /// <summary>
    /// 添加两期移动平均趋势线
    /// </summary>
    msoElementTrendlineAddTwoPeriodMovingAverage = 604,

    // 误差线设置 (700-703)
    /// <summary>
    /// 不显示误差线
    /// </summary>
    msoElementErrorBarNone = 700,
    /// <summary>
    /// 显示标准误差误差线
    /// </summary>
    msoElementErrorBarStandardError = 701,
    /// <summary>
    /// 显示百分比误差线
    /// </summary>
    msoElementErrorBarPercentage = 702,
    /// <summary>
    /// 显示标准偏差误差线
    /// </summary>
    msoElementErrorBarStandardDeviation = 703,

    // 线条设置 (800-804)
    /// <summary>
    /// 不显示线条
    /// </summary>
    msoElementLineNone = 800,
    /// <summary>
    /// 显示垂直线
    /// </summary>
    msoElementLineDropLine = 801,
    /// <summary>
    /// 显示高低点连线
    /// </summary>
    msoElementLineHiLoLine = 802,
    /// <summary>
    /// 显示系列线
    /// </summary>
    msoElementLineSeriesLine = 803,
    /// <summary>
    /// 显示垂直和高低点连线
    /// </summary>
    msoElementLineDropHiLoLine = 804,

    // 上下限柱设置 (900-901)
    /// <summary>
    /// 不显示上下限柱
    /// </summary>
    msoElementUpDownBarsNone = 900,
    /// <summary>
    /// 显示上下限柱
    /// </summary>
    msoElementUpDownBarsShow = 901,

    // 绘图区设置 (1000-1001)
    /// <summary>
    /// 不显示绘图区
    /// </summary>
    msoElementPlotAreaNone = 1000,
    /// <summary>
    /// 显示绘图区
    /// </summary>
    msoElementPlotAreaShow = 1001,

    // 图表墙设置 (1100-1101)
    /// <summary>
    /// 不显示图表墙
    /// </summary>
    msoElementChartWallNone = 1100,
    /// <summary>
    /// 显示图表墙
    /// </summary>
    msoElementChartWallShow = 1101,

    // 图表底板设置 (1200-1201)
    /// <summary>
    /// 不显示图表底板
    /// </summary>
    msoElementChartFloorNone = 1200,
    /// <summary>
    /// 显示图表底板
    /// </summary>
    msoElementChartFloorShow = 1201
}