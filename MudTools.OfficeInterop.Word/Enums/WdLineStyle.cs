namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定线条的样式类型，用于文档中的边框和线条格式化
/// </summary>
public enum WdLineStyle
{
    /// <summary>
    /// 无线条样式
    /// </summary>
    wdLineStyleNone,
    /// <summary>
    /// 单线样式
    /// </summary>
    wdLineStyleSingle,
    /// <summary>
    /// 点线样式
    /// </summary>
    wdLineStyleDot,
    /// <summary>
    /// 小间隔虚线样式
    /// </summary>
    wdLineStyleDashSmallGap,
    /// <summary>
    /// 大间隔虚线样式
    /// </summary>
    wdLineStyleDashLargeGap,
    /// <summary>
    /// 一线一点样式
    /// </summary>
    wdLineStyleDashDot,
    /// <summary>
    /// 一线两点样式
    /// </summary>
    wdLineStyleDashDotDot,
    /// <summary>
    /// 双线样式
    /// </summary>
    wdLineStyleDouble,
    /// <summary>
    /// 三线样式
    /// </summary>
    wdLineStyleTriple,
    /// <summary>
    /// 细线-粗线-细线(小间隔)样式
    /// </summary>
    wdLineStyleThinThickSmallGap,
    /// <summary>
    /// 粗线-细线-粗线(小间隔)样式
    /// </summary>
    wdLineStyleThickThinSmallGap,
    /// <summary>
    /// 细线-粗线-细线(小间隔)样式
    /// </summary>
    wdLineStyleThinThickThinSmallGap,
    /// <summary>
    /// 细线-粗线(中等间隔)样式
    /// </summary>
    wdLineStyleThinThickMedGap,
    /// <summary>
    /// 粗线-细线(中等间隔)样式
    /// </summary>
    wdLineStyleThickThinMedGap,
    /// <summary>
    /// 细线-粗线-细线(中等间隔)样式
    /// </summary>
    wdLineStyleThinThickThinMedGap,
    /// <summary>
    /// 细线-粗线(大间隔)样式
    /// </summary>
    wdLineStyleThinThickLargeGap,
    /// <summary>
    /// 粗线-细线(大间隔)样式
    /// </summary>
    wdLineStyleThickThinLargeGap,
    /// <summary>
    /// 细线-粗线-细线(大间隔)样式
    /// </summary>
    wdLineStyleThinThickThinLargeGap,
    /// <summary>
    /// 单波浪线样式
    /// </summary>
    wdLineStyleSingleWavy,
    /// <summary>
    /// 双波浪线样式
    /// </summary>
    wdLineStyleDoubleWavy,
    /// <summary>
    /// 交替点划线样式
    /// </summary>
    wdLineStyleDashDotStroked,
    /// <summary>
    /// 凸起3D样式
    /// </summary>
    wdLineStyleEmboss3D,
    /// <summary>
    /// 凹陷3D样式
    /// </summary>
    wdLineStyleEngrave3D,
    /// <summary>
    /// 外凸样式
    /// </summary>
    wdLineStyleOutset,
    /// <summary>
    /// 内嵌样式
    /// </summary>
    wdLineStyleInset
}