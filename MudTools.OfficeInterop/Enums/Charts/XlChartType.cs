//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定图表类型
/// </summary>
[TypeLibType(16)]
public enum XlChartType
{
    /// <summary>
    /// 三维簇状柱形图
    /// </summary>
    xlColumnClustered = 51,

    /// <summary>
    /// 堆积柱形图
    /// </summary>
    xlColumnStacked = 52,

    /// <summary>
    /// 百分比堆积柱形图
    /// </summary>
    xlColumnStacked100 = 53,

    /// <summary>
    /// 三维簇状柱形图
    /// </summary>
    xl3DColumnClustered = 54,

    /// <summary>
    /// 三维堆积柱形图
    /// </summary>
    xl3DColumnStacked = 55,

    /// <summary>
    /// 三维百分比堆积条形图
    /// </summary>
    xl3DColumnStacked100 = 56,

    /// <summary>
    /// 簇状条形图
    /// </summary>
    xlBarClustered = 57,

    /// <summary>
    /// 堆积条形图
    /// </summary>
    xlBarStacked = 58,

    /// <summary>
    /// 百分比堆积条形图
    /// </summary>
    xlBarStacked100 = 59,

    /// <summary>
    /// 三维簇状条形图
    /// </summary>
    xl3DBarClustered = 60,

    /// <summary>
    /// 三维堆积条形图
    /// </summary>
    xl3DBarStacked = 61,

    /// <summary>
    /// 三维百分比堆积条形图
    /// </summary>
    xl3DBarStacked100 = 62,

    /// <summary>
    /// 堆积折线图
    /// </summary>
    xlLineStacked = 63,

    /// <summary>
    /// 百分比堆积折线图
    /// </summary>
    xlLineStacked100 = 64,

    /// <summary>
    /// 带数据标记的折线图
    /// </summary>
    xlLineMarkers = 65,

    /// <summary>
    /// 带数据标记的堆积折线图
    /// </summary>
    xlLineMarkersStacked = 66,

    /// <summary>
    /// 带数据标记的百分比堆积折线图
    /// </summary>
    xlLineMarkersStacked100 = 67,

    /// <summary>
    /// 复合饼图
    /// </summary>
    xlPieOfPie = 68,

    /// <summary>
    /// 分离型饼图
    /// </summary>
    xlPieExploded = 69,

    /// <summary>
    /// 分离型三维饼图
    /// </summary>
    xl3DPieExploded = 70,

    /// <summary>
    /// 复合条饼图
    /// </summary>
    xlBarOfPie = 71,

    /// <summary>
    /// 带平滑线的散点图
    /// </summary>
    xlXYScatterSmooth = 72,

    /// <summary>
    /// 带平滑线无数据标记的散点图
    /// </summary>
    xlXYScatterSmoothNoMarkers = 73,

    /// <summary>
    /// 带直线的散点图
    /// </summary>
    xlXYScatterLines = 74,

    /// <summary>
    /// 带直线无数据标记的散点图
    /// </summary>
    xlXYScatterLinesNoMarkers = 75,

    /// <summary>
    /// 堆积面积图
    /// </summary>
    xlAreaStacked = 76,

    /// <summary>
    /// 百分比堆积面积图
    /// </summary>
    xlAreaStacked100 = 77,

    /// <summary>
    /// 三维堆积面积图
    /// </summary>
    xl3DAreaStacked = 78,

    /// <summary>
    /// 百分比堆积面积图
    /// </summary>
    xl3DAreaStacked100 = 79,

    /// <summary>
    /// 分离型圆环图
    /// </summary>
    xlDoughnutExploded = 80,

    /// <summary>
    /// 带数据标记的雷达图
    /// </summary>
    xlRadarMarkers = 81,

    /// <summary>
    /// 填充雷达图
    /// </summary>
    xlRadarFilled = 82,

    /// <summary>
    /// 三维曲面图
    /// </summary>
    xlSurface = 83,

    /// <summary>
    /// 三维曲面图（框架）
    /// </summary>
    xlSurfaceWireframe = 84,

    /// <summary>
    /// 曲面图（俯视）
    /// </summary>
    xlSurfaceTopView = 85,

    /// <summary>
    /// 曲面图（俯视框架）
    /// </summary>
    xlSurfaceTopViewWireframe = 86,

    /// <summary>
    /// 气泡图
    /// </summary>
    xlBubble = 15,

    /// <summary>
    /// 三维效果气泡图
    /// </summary>
    xlBubble3DEffect = 87,

    /// <summary>
    /// 盘高-盘低-收盘图
    /// </summary>
    xlStockHLC = 88,

    /// <summary>
    /// 开盘-盘高-盘低-收盘图
    /// </summary>
    xlStockOHLC = 89,

    /// <summary>
    /// 成交量-盘高-盘低-收盘图
    /// </summary>
    xlStockVHLC = 90,

    /// <summary>
    /// 成交量-开盘-盘高-盘低-收盘图
    /// </summary>
    xlStockVOHLC = 91,

    /// <summary>
    /// 簇状圆锥柱形图
    /// </summary>
    xlCylinderColClustered = 92,

    /// <summary>
    /// 堆积圆柱条形图
    /// </summary>
    xlCylinderColStacked = 93,

    /// <summary>
    /// 百分比堆积圆柱柱形图
    /// </summary>
    xlCylinderColStacked100 = 94,

    /// <summary>
    /// 簇状圆柱条形图
    /// </summary>
    xlCylinderBarClustered = 95,

    /// <summary>
    /// 堆积圆柱条形图
    /// </summary>
    xlCylinderBarStacked = 96,

    /// <summary>
    /// 百分比堆积圆柱条形图
    /// </summary>
    xlCylinderBarStacked100 = 97,

    /// <summary>
    /// 三维圆柱柱形图
    /// </summary>
    xlCylinderCol = 98,

    /// <summary>
    /// 簇状圆锥柱形图
    /// </summary>
    xlConeColClustered = 99,

    /// <summary>
    /// 堆积圆锥柱形图
    /// </summary>
    xlConeColStacked = 100,

    /// <summary>
    /// 百分比堆积圆锥柱形图
    /// </summary>
    xlConeColStacked100 = 101,

    /// <summary>
    /// 簇状圆锥条形图
    /// </summary>
    xlConeBarClustered = 102,

    /// <summary>
    /// 堆积圆锥条形图
    /// </summary>
    xlConeBarStacked = 103,

    /// <summary>
    /// 百分比堆积圆锥条形图
    /// </summary>
    xlConeBarStacked100 = 104,

    /// <summary>
    /// 三维圆锥柱形图
    /// </summary>
    xlConeCol = 105,

    /// <summary>
    /// 簇状棱锥柱形图
    /// </summary>
    xlPyramidColClustered = 106,

    /// <summary>
    /// 堆积棱锥柱形图
    /// </summary>
    xlPyramidColStacked = 107,

    /// <summary>
    /// 百分比堆积棱锥柱形图
    /// </summary>
    xlPyramidColStacked100 = 108,

    /// <summary>
    /// 簇状棱锥条形图
    /// </summary>
    xlPyramidBarClustered = 109,

    /// <summary>
    /// 堆积棱锥条形图
    /// </summary>
    xlPyramidBarStacked = 110,

    /// <summary>
    /// 百分比堆积棱锥条形图
    /// </summary>
    xlPyramidBarStacked100 = 111,

    /// <summary>
    /// 三维棱锥柱形图
    /// </summary>
    xlPyramidCol = 112,

    /// <summary>
    /// 三维柱形图
    /// </summary>
    xl3DColumn = -4100,

    /// <summary>
    /// 折线图
    /// </summary>
    xlLine = 4,

    /// <summary>
    /// 三维折线图
    /// </summary>
    xl3DLine = -4101,

    /// <summary>
    /// 三维饼图
    /// </summary>
    xl3DPie = -4102,

    /// <summary>
    /// 饼图
    /// </summary>
    xlPie = 5,

    /// <summary>
    /// 散点图
    /// </summary>
    xlXYScatter = -4169,

    /// <summary>
    /// 三维面积图
    /// </summary>
    xl3DArea = -4098,

    /// <summary>
    /// 面积图
    /// </summary>
    xlArea = 1,

    /// <summary>
    /// 圆环图
    /// </summary>
    xlDoughnut = -4120,

    /// <summary>
    /// 雷达图
    /// </summary>
    xlRadar = -4151,

    /// <summary>
    /// 组合图
    /// </summary>
    xlCombo = -4152,

    /// <summary>
    /// 簇状柱形图-折线图组合
    /// </summary>
    xlComboColumnClusteredLine = 113,

    /// <summary>
    /// 簇状柱形图-折线图组合（带次坐标轴）
    /// </summary>
    xlComboColumnClusteredLineSecondaryAxis = 114,

    /// <summary>
    /// 堆积面积图-簇状柱形图组合
    /// </summary>
    xlComboAreaStackedColumnClustered = 115,

    /// <summary>
    /// 其他组合图表
    /// </summary>
    xlOtherCombinations = 116,

    /// <summary>
    /// 推荐的图表类型
    /// </summary>
    xlSuggestedChart = -2
}