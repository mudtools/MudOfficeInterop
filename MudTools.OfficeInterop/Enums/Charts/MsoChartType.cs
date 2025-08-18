//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 图表类型枚举
/// 用于指定Excel中可用的各种图表类型
/// </summary>
public enum MsoChartType
{
    /// <summary>
    /// 柱形图-簇状柱形图
    /// 以簇状形式显示数据的柱形图
    /// </summary>
    xlColumnClustered = 51,

    /// <summary>
    /// 柱形图-堆积柱形图
    /// 以堆积形式显示数据的柱形图
    /// </summary>
    xlColumnStacked = 52,

    /// <summary>
    /// 柱形图-百分比堆积柱形图
    /// 以百分比堆积形式显示数据的柱形图
    /// </summary>
    xlColumnStacked100 = 53,

    /// <summary>
    /// 三维柱形图-簇状柱形图
    /// 以三维形式显示的簇状柱形图
    /// </summary>
    xl3DColumnClustered = 54,

    /// <summary>
    /// 三维柱形图-堆积柱形图
    /// 以三维形式显示的堆积柱形图
    /// </summary>
    xl3DColumnStacked = 55,

    /// <summary>
    /// 三维柱形图-百分比堆积柱形图
    /// 以三维形式显示的百分比堆积柱形图
    /// </summary>
    xl3DColumnStacked100 = 56,

    /// <summary>
    /// 条形图-簇状条形图
    /// 以簇状形式显示数据的条形图（水平柱形图）
    /// </summary>
    xlBarClustered = 57,

    /// <summary>
    /// 条形图-堆积条形图
    /// 以堆积形式显示数据的条形图（水平柱形图）
    /// </summary>
    xlBarStacked = 58,

    /// <summary>
    /// 条形图-百分比堆积条形图
    /// 以百分比堆积形式显示数据的条形图（水平柱形图）
    /// </summary>
    xlBarStacked100 = 59,

    /// <summary>
    /// 三维条形图-簇状条形图
    /// 以三维形式显示的簇状条形图
    /// </summary>
    xl3DBarClustered = 60,

    /// <summary>
    /// 三维条形图-堆积条形图
    /// 以三维形式显示的堆积条形图
    /// </summary>
    xl3DBarStacked = 61,

    /// <summary>
    /// 三维条形图-百分比堆积条形图
    /// 以三维形式显示的百分比堆积条形图
    /// </summary>
    xl3DBarStacked100 = 62,

    /// <summary>
    /// 折线图-堆积折线图
    /// 数据点以堆积方式显示的折线图
    /// </summary>
    xlLineStacked = 63,

    /// <summary>
    /// 折线图-百分比堆积折线图
    /// 数据点以百分比堆积方式显示的折线图
    /// </summary>
    xlLineStacked100 = 64,

    /// <summary>
    /// 折线图-带数据标记的折线图
    /// 带有数据标记的折线图
    /// </summary>
    xlLineMarkers = 65,

    /// <summary>
    /// 折线图-带数据标记的堆积折线图
    /// 带有数据标记的堆积折线图
    /// </summary>
    xlLineMarkersStacked = 66,

    /// <summary>
    /// 折线图-带数据标记的百分比堆积折线图
    /// 带有数据标记的百分比堆积折线图
    /// </summary>
    xlLineMarkersStacked100 = 67,

    /// <summary>
    /// 饼图-复合饼图
    /// 将较小的数据点合并到第二个饼图中的饼图
    /// </summary>
    xlPieOfPie = 68,

    /// <summary>
    /// 饼图-分离型饼图
    /// 部分扇区分离的饼图
    /// </summary>
    xlPieExploded = 69,

    /// <summary>
    /// 三维饼图-分离型三维饼图
    /// 部分扇区分离的三维饼图
    /// </summary>
    xl3DPieExploded = 70,

    /// <summary>
    /// 饼图-复合条饼图
    /// 将较小的数据点合并到第二个条形图中的饼图
    /// </summary>
    xlBarOfPie = 71,

    /// <summary>
    /// 散点图-平滑线散点图
    /// 用平滑曲线连接数据点的散点图
    /// </summary>
    xlXYScatterSmooth = 72,

    /// <summary>
    /// 散点图-无数据标记的平滑线散点图
    /// 用平滑曲线连接数据点但不显示数据标记的散点图
    /// </summary>
    xlXYScatterSmoothNoMarkers = 73,

    /// <summary>
    /// 散点图-折线散点图
    /// 用直线连接数据点的散点图
    /// </summary>
    xlXYScatterLines = 74,

    /// <summary>
    /// 散点图-无数据标记的折线散点图
    /// 用直线连接数据点但不显示数据标记的散点图
    /// </summary>
    xlXYScatterLinesNoMarkers = 75,

    /// <summary>
    /// 面积图-堆积面积图
    /// 以堆积形式显示数据的面积图
    /// </summary>
    xlAreaStacked = 76,

    /// <summary>
    /// 面积图-百分比堆积面积图
    /// 以百分比堆积形式显示数据的面积图
    /// </summary>
    xlAreaStacked100 = 77,

    /// <summary>
    /// 三维面积图-堆积面积图
    /// 以三维形式显示的堆积面积图
    /// </summary>
    xl3DAreaStacked = 78,

    /// <summary>
    /// 三维面积图-百分比堆积面积图
    /// 以三维形式显示的百分比堆积面积图
    /// </summary>
    xl3DAreaStacked100 = 79,

    /// <summary>
    /// 圆环图-分离型圆环图
    /// 部分扇区分离开的圆环图
    /// </summary>
    xlDoughnutExploded = 80,

    /// <summary>
    /// 雷达图-带数据标记的雷达图
    /// 带有数据标记的雷达图
    /// </summary>
    xlRadarMarkers = 81,

    /// <summary>
    /// 雷达图-填充雷达图
    /// 填充区域的雷达图
    /// </summary>
    xlRadarFilled = 82,

    /// <summary>
    /// 曲面图-三维曲面图
    /// 显示数据在三维空间中的曲面图
    /// </summary>
    xlSurface = 83,

    /// <summary>
    /// 曲面图-三维曲面图（网格线）
    /// 显示网格线的三维曲面图
    /// </summary>
    xlSurfaceWireframe = 84,

    /// <summary>
    /// 曲面图-俯视图
    /// 从上方观察的曲面图
    /// </summary>
    xlSurfaceTopView = 85,

    /// <summary>
    /// 曲面图-俯视图（网格线）
    /// 从上方观察并显示网格线的曲面图
    /// </summary>
    xlSurfaceTopViewWireframe = 86,

    /// <summary>
    /// 气泡图
    /// 用气泡大小表示第三维数据的散点图
    /// </summary>
    xlBubble = 15,

    /// <summary>
    /// 气泡图-三维气泡图效果
    /// 具有三维效果的气泡图
    /// </summary>
    xlBubble3DEffect = 87,

    /// <summary>
    /// 股价图-最高价-最低价-收盘价图
    /// 显示股票最高价、最低价和收盘价的图表
    /// </summary>
    xlStockHLC = 88,

    /// <summary>
    /// 股价图-开盘价-最高价-最低价-收盘价图
    /// 显示股票开盘价、最高价、最低价和收盘价的图表
    /// </summary>
    xlStockOHLC = 89,

    /// <summary>
    /// 股价图-最高价-最低价-收盘价-成交量图
    /// 显示股票最高价、最低价、收盘价和成交量的图表
    /// </summary>
    xlStockVHLC = 90,

    /// <summary>
    /// 股价图-开盘价-最高价-最低价-收盘价-成交量图
    /// 显示股票开盘价、最高价、最低价、收盘价和成交量的图表
    /// </summary>
    xlStockVOHLC = 91,

    /// <summary>
    /// 圆柱图-簇状圆柱图
    /// 以簇状形式显示的圆柱图
    /// </summary>
    xlCylinderColClustered = 92,

    /// <summary>
    /// 圆柱图-堆积圆柱图
    /// 以堆积形式显示的圆柱图
    /// </summary>
    xlCylinderColStacked = 93,

    /// <summary>
    /// 圆柱图-百分比堆积圆柱图
    /// 以百分比堆积形式显示的圆柱图
    /// </summary>
    xlCylinderColStacked100 = 94,

    /// <summary>
    /// 圆柱图-簇状圆柱条形图
    /// 以簇状形式显示的水平圆柱图
    /// </summary>
    xlCylinderBarClustered = 95,

    /// <summary>
    /// 圆柱图-堆积圆柱条形图
    /// 以堆积形式显示的水平圆柱图
    /// </summary>
    xlCylinderBarStacked = 96,

    /// <summary>
    /// 圆柱图-百分比堆积圆柱条形图
    /// 以百分比堆积形式显示的水平圆柱图
    /// </summary>
    xlCylinderBarStacked100 = 97,

    /// <summary>
    /// 圆柱图-圆柱图
    /// 标准的圆柱图
    /// </summary>
    xlCylinderCol = 98,

    /// <summary>
    /// 圆锥图-簇状圆锥图
    /// 以簇状形式显示的圆锥图
    /// </summary>
    xlConeColClustered = 99,

    /// <summary>
    /// 圆锥图-堆积圆锥图
    /// 以堆积形式显示的圆锥图
    /// </summary>
    xlConeColStacked = 100,

    /// <summary>
    /// 圆锥图-百分比堆积圆锥图
    /// 以百分比堆积形式显示的圆锥图
    /// </summary>
    xlConeColStacked100 = 101,

    /// <summary>
    /// 圆锥图-簇状圆锥条形图
    /// 以簇状形式显示的水平圆锥图
    /// </summary>
    xlConeBarClustered = 102,

    /// <summary>
    /// 圆锥图-堆积圆锥条形图
    /// 以堆积形式显示的水平圆锥图
    /// </summary>
    xlConeBarStacked = 103,

    /// <summary>
    /// 圆锥图-百分比堆积圆锥条形图
    /// 以百分比堆积形式显示的水平圆锥图
    /// </summary>
    xlConeBarStacked100 = 104,

    /// <summary>
    /// 圆锥图-圆锥图
    /// 标准的圆锥图
    /// </summary>
    xlConeCol = 105,

    /// <summary>
    /// 棱锥图-簇状棱锥图
    /// 以簇状形式显示的棱锥图
    /// </summary>
    xlPyramidColClustered = 106,

    /// <summary>
    /// 棱锥图-堆积棱锥图
    /// 以堆积形式显示的棱锥图
    /// </summary>
    xlPyramidColStacked = 107,

    /// <summary>
    /// 棱锥图-百分比堆积棱锥图
    /// 以百分比堆积形式显示的棱锥图
    /// </summary>
    xlPyramidColStacked100 = 108,

    /// <summary>
    /// 棱锥图-簇状棱锥条形图
    /// 以簇状形式显示的水平棱锥图
    /// </summary>
    xlPyramidBarClustered = 109,

    /// <summary>
    /// 棱锥图-堆积棱锥条形图
    /// 以堆积形式显示的水平棱锥图
    /// </summary>
    xlPyramidBarStacked = 110,

    /// <summary>
    /// 棱锥图-百分比堆积棱锥条形图
    /// 以百分比堆积形式显示的水平棱锥图
    /// </summary>
    xlPyramidBarStacked100 = 111,

    /// <summary>
    /// 棱锥图-棱锥图
    /// 标准的棱锥图
    /// </summary>
    xlPyramidCol = 112,

    /// <summary>
    /// 三维柱形图
    /// 标准的三维柱形图
    /// </summary>
    xl3DColumn = -4100,

    /// <summary>
    /// 折线图
    /// 标准的折线图
    /// </summary>
    xlLine = 4,

    /// <summary>
    /// 三维折线图
    /// 标准的三维折线图
    /// </summary>
    xl3DLine = -4101,

    /// <summary>
    /// 三维饼图
    /// 标准的三维饼图
    /// </summary>
    xl3DPie = -4102,

    /// <summary>
    /// 饼图
    /// 标准的饼图
    /// </summary>
    xlPie = 5,

    /// <summary>
    /// 散点图
    /// 标准的散点图
    /// </summary>
    xlXYScatter = -4169,

    /// <summary>
    /// 三维面积图
    /// 标准的三维面积图
    /// </summary>
    xl3DArea = -4098,

    /// <summary>
    /// 面积图
    /// 标准的面积图
    /// </summary>
    xlArea = 1,

    /// <summary>
    /// 圆环图
    /// 标准的圆环图
    /// </summary>
    xlDoughnut = -4120,

    /// <summary>
    /// 雷达图
    /// 标准的雷达图
    /// </summary>
    xlRadar = -4151
}