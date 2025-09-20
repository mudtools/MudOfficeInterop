namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定图表类型
/// </summary>
public enum XlChartType
{
    /// <summary>
    /// 簇状柱形图
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
    /// 三维百分比堆积柱形图
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
    /// 爆炸饼图
    /// </summary>
    xlPieExploded = 69,
    /// <summary>
    /// 三维爆炸饼图
    /// </summary>
    xl3DPieExploded = 70,
    /// <summary>
    /// 复合条饼图
    /// </summary>
    xlBarOfPie = 71,
    /// <summary>
    /// 平滑散点图
    /// </summary>
    xlXYScatterSmooth = 72,
    /// <summary>
    /// 无数据标记的平滑散点图
    /// </summary>
    xlXYScatterSmoothNoMarkers = 73,
    /// <summary>
    /// 折线散点图
    /// </summary>
    xlXYScatterLines = 74,
    /// <summary>
    /// 无数据标记的折线散点图
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
    /// 三维百分比堆积面积图
    /// </summary>
    xl3DAreaStacked100 = 79,
    /// <summary>
    /// 爆炸圆环图
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
    /// 三维曲面图（框架图）
    /// </summary>
    xlSurfaceWireframe = 84,
    /// <summary>
    /// 俯视三维曲面图
    /// </summary>
    xlSurfaceTopView = 85,
    /// <summary>
    /// 俯视三维曲面图（框架图）
    /// </summary>
    xlSurfaceTopViewWireframe = 86,
    /// <summary>
    /// 气泡图
    /// </summary>
    xlBubble = 15,
    /// <summary>
    /// 三维气泡图
    /// </summary>
    xlBubble3DEffect = 87,
    /// <summary>
    /// 股价图（高-低-收盘）
    /// </summary>
    xlStockHLC = 88,
    /// <summary>
    /// 股价图（开盘-高-低-收盘）
    /// </summary>
    xlStockOHLC = 89,
    /// <summary>
    /// 股价图（成交量-高-低-收盘）
    /// </summary>
    xlStockVHLC = 90,
    /// <summary>
    /// 股价图（成交量-开盘-高-低-收盘）
    /// </summary>
    xlStockVOHLC = 91,
    /// <summary>
    /// 簇状圆柱图
    /// </summary>
    xlCylinderColClustered = 92,
    /// <summary>
    /// 堆积圆柱图
    /// </summary>
    xlCylinderColStacked = 93,
    /// <summary>
    /// 百分比堆积圆柱图
    /// </summary>
    xlCylinderColStacked100 = 94,
    /// <summary>
    /// 簇状水平圆柱图
    /// </summary>
    xlCylinderBarClustered = 95,
    /// <summary>
    /// 堆积水平圆柱图
    /// </summary>
    xlCylinderBarStacked = 96,
    /// <summary>
    /// 百分比堆积水平圆柱图
    /// </summary>
    xlCylinderBarStacked100 = 97,
    /// <summary>
    /// 圆柱图
    /// </summary>
    xlCylinderCol = 98,
    /// <summary>
    /// 簇状圆锥图
    /// </summary>
    xlConeColClustered = 99,
    /// <summary>
    /// 堆积圆锥图
    /// </summary>
    xlConeColStacked = 100,
    /// <summary>
    /// 百分比堆积圆锥图
    /// </summary>
    xlConeColStacked100 = 101,
    /// <summary>
    /// 簇状水平圆锥图
    /// </summary>
    xlConeBarClustered = 102,
    /// <summary>
    /// 堆积水平圆锥图
    /// </summary>
    xlConeBarStacked = 103,
    /// <summary>
    /// 百分比堆积水平圆锥图
    /// </summary>
    xlConeBarStacked100 = 104,
    /// <summary>
    /// 圆锥图
    /// </summary>
    xlConeCol = 105,
    /// <summary>
    /// 簇状棱锥图
    /// </summary>
    xlPyramidColClustered = 106,
    /// <summary>
    /// 堆积棱锥图
    /// </summary>
    xlPyramidColStacked = 107,
    /// <summary>
    /// 百分比堆积棱锥图
    /// </summary>
    xlPyramidColStacked100 = 108,
    /// <summary>
    /// 簇状水平棱锥图
    /// </summary>
    xlPyramidBarClustered = 109,
    /// <summary>
    /// 堆积水平棱锥图
    /// </summary>
    xlPyramidBarStacked = 110,
    /// <summary>
    /// 百分比堆积水平棱锥图
    /// </summary>
    xlPyramidBarStacked100 = 111,
    /// <summary>
    /// 棱锥图
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
    xlRadar = -4151
}