namespace MudTools.OfficeInterop;

/// <summary>
/// 指定图表数据标签的显示类型
/// </summary>
public enum XlDataLabelsType
{
    /// <summary>
    /// 不显示数据标签
    /// </summary>
    xlDataLabelsShowNone = -4142,

    /// <summary>
    /// 显示数值
    /// </summary>
    xlDataLabelsShowValue = 2,

    /// <summary>
    /// 显示百分比
    /// </summary>
    xlDataLabelsShowPercent = 3,

    /// <summary>
    /// 显示类别名称标签
    /// </summary>
    xlDataLabelsShowLabel = 4,

    /// <summary>
    /// 同时显示标签和百分比
    /// </summary>
    xlDataLabelsShowLabelAndPercent = 5,

    /// <summary>
    /// 显示气泡大小
    /// </summary>
    xlDataLabelsShowBubbleSizes = 6
}