namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定在页面上的垂直位置相对于什么元素进行定位
/// </summary>
public enum WdRelativeVerticalPosition
{
    /// <summary>
/// 相对于页边距定位
/// </summary>
    wdRelativeVerticalPositionMargin,
    /// <summary>
/// 相对于页面定位
/// </summary>
    wdRelativeVerticalPositionPage,
    /// <summary>
/// 相对于段落定位
/// </summary>
    wdRelativeVerticalPositionParagraph,
    /// <summary>
/// 相对于行定位
/// </summary>
    wdRelativeVerticalPositionLine,
    /// <summary>
/// 相对于上边距区域定位
/// </summary>
    wdRelativeVerticalPositionTopMarginArea,
    /// <summary>
/// 相对于下边距区域定位
/// </summary>
    wdRelativeVerticalPositionBottomMarginArea,
    /// <summary>
/// 相对于内边距区域定位
/// </summary>
    wdRelativeVerticalPositionInnerMarginArea,
    /// <summary>
/// 相对于外边距区域定位
/// </summary>
    wdRelativeVerticalPositionOuterMarginArea
}