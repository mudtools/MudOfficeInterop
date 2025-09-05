namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定在页面上的水平位置相对于什么元素进行定位
/// </summary>
public enum WdRelativeHorizontalPosition
{
    /// <summary>
    /// 相对于页边距定位
    /// </summary>
    wdRelativeHorizontalPositionMargin,
    /// <summary>
    /// 相对于页面定位
    /// </summary>
    wdRelativeHorizontalPositionPage,
    /// <summary>
    /// 相对于栏定位
    /// </summary>
    wdRelativeHorizontalPositionColumn,
    /// <summary>
    /// 相对于字符定位
    /// </summary>
    wdRelativeHorizontalPositionCharacter,
    /// <summary>
    /// 相对于左边距区域定位
    /// </summary>
    wdRelativeHorizontalPositionLeftMarginArea,
    /// <summary>
    /// 相对于右边距区域定位
    /// </summary>
    wdRelativeHorizontalPositionRightMarginArea,
    /// <summary>
    /// 相对于内边距区域定位
    /// </summary>
    wdRelativeHorizontalPositionInnerMarginArea,
    /// <summary>
    /// 相对于外边距区域定位
    /// </summary>
    wdRelativeHorizontalPositionOuterMarginArea
}