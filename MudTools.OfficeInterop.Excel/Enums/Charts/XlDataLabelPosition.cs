namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 指定图表数据标签的位置
/// </summary>
public enum XlDataLabelPosition
{
    /// <summary>
    /// 将数据标签居中放置在数据点上
    /// </summary>
    xlLabelPositionCenter = -4108,

    /// <summary>
    /// 将数据标签放置在数据点上方
    /// </summary>
    xlLabelPositionAbove = 0,

    /// <summary>
    /// 将数据标签放置在数据点下方
    /// </summary>
    xlLabelPositionBelow = 1,

    /// <summary>
    /// 将数据标签放置在数据点左侧
    /// </summary>
    xlLabelPositionLeft = -4131,

    /// <summary>
    /// 将数据标签放置在数据点右侧
    /// </summary>
    xlLabelPositionRight = -4152,

    /// <summary>
    /// 将数据标签放置在数据点外部末端
    /// </summary>
    xlLabelPositionOutsideEnd = 2,

    /// <summary>
    /// 将数据标签放置在数据点内部末端
    /// </summary>
    xlLabelPositionInsideEnd = 3,

    /// <summary>
    /// 将数据标签放置在数据点内部基部
    /// </summary>
    xlLabelPositionInsideBase = 4,

    /// <summary>
    /// 自动选择最适合的位置放置数据标签
    /// </summary>
    xlLabelPositionBestFit = 5,

    /// <summary>
    /// 数据标签位置混合模式
    /// </summary>
    xlLabelPositionMixed = 6,

    /// <summary>
    /// 自定义数据标签位置
    /// </summary>
    xlLabelPositionCustom = 7
}