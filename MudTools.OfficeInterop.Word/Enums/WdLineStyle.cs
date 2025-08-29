namespace MudTools.OfficeInterop.Word;

/// &lt;summary&gt;
/// 指定线条的样式类型，用于文档中的边框和线条格式化
/// &lt;/summary&gt;
public enum WdLineStyle
{
    /// &lt;summary&gt;
    /// 无线条样式
    /// &lt;/summary&gt;
    wdLineStyleNone,
    /// &lt;summary&gt;
    /// 单线样式
    /// &lt;/summary&gt;
    wdLineStyleSingle,
    /// &lt;summary&gt;
    /// 点线样式
    /// &lt;/summary&gt;
    wdLineStyleDot,
    /// &lt;summary&gt;
    /// 小间隔虚线样式
    /// &lt;/summary&gt;
    wdLineStyleDashSmallGap,
    /// &lt;summary&gt;
    /// 大间隔虚线样式
    /// &lt;/summary&gt;
    wdLineStyleDashLargeGap,
    /// &lt;summary&gt;
    /// 一线一点样式
    /// &lt;/summary&gt;
    wdLineStyleDashDot,
    /// &lt;summary&gt;
    /// 一线两点样式
    /// &lt;/summary&gt;
    wdLineStyleDashDotDot,
    /// &lt;summary&gt;
    /// 双线样式
    /// &lt;/summary&gt;
    wdLineStyleDouble,
    /// &lt;summary&gt;
    /// 三线样式
    /// &lt;/summary&gt;
    wdLineStyleTriple,
    /// &lt;summary&gt;
    /// 细线-粗线-细线(小间隔)样式
    /// &lt;/summary&gt;
    wdLineStyleThinThickSmallGap,
    /// &lt;summary&gt;
    /// 粗线-细线-粗线(小间隔)样式
    /// &lt;/summary&gt;
    wdLineStyleThickThinSmallGap,
    /// &lt;summary&gt;
    /// 细线-粗线-细线(小间隔)样式
    /// &lt;/summary&gt;
    wdLineStyleThinThickThinSmallGap,
    /// &lt;summary&gt;
    /// 细线-粗线(中等间隔)样式
    /// &lt;/summary&gt;
    wdLineStyleThinThickMedGap,
    /// &lt;summary&gt;
    /// 粗线-细线(中等间隔)样式
    /// &lt;/summary&gt;
    wdLineStyleThickThinMedGap,
    /// &lt;summary&gt;
    /// 细线-粗线-细线(中等间隔)样式
    /// &lt;/summary&gt;
    wdLineStyleThinThickThinMedGap,
    /// &lt;summary&gt;
    /// 细线-粗线(大间隔)样式
    /// &lt;/summary&gt;
    wdLineStyleThinThickLargeGap,
    /// &lt;summary&gt;
    /// 粗线-细线(大间隔)样式
    /// &lt;/summary&gt;
    wdLineStyleThickThinLargeGap,
    /// &lt;summary&gt;
    /// 细线-粗线-细线(大间隔)样式
    /// &lt;/summary&gt;
    wdLineStyleThinThickThinLargeGap,
    /// &lt;summary&gt;
    /// 单波浪线样式
    /// &lt;/summary&gt;
    wdLineStyleSingleWavy,
    /// &lt;summary&gt;
    /// 双波浪线样式
    /// &lt;/summary&gt;
    wdLineStyleDoubleWavy,
    /// &lt;summary&gt;
    /// 交替点划线样式
    /// &lt;/summary&gt;
    wdLineStyleDashDotStroked,
    /// &lt;summary&gt;
    /// 凸起3D样式
    /// &lt;/summary&gt;
    wdLineStyleEmboss3D,
    /// &lt;summary&gt;
    /// 凹陷3D样式
    /// &lt;/summary&gt;
    wdLineStyleEngrave3D,
    /// &lt;summary&gt;
    /// 外凸样式
    /// &lt;/summary&gt;
    wdLineStyleOutset,
    /// &lt;summary&gt;
    /// 内嵌样式
    /// &lt;/summary&gt;
    wdLineStyleInset
}