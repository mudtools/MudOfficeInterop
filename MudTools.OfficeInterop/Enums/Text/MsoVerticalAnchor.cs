namespace MudTools.OfficeInterop;

/// <summary>
/// 指定文本框内文本的垂直对齐方式
/// </summary>
public enum MsoVerticalAnchor
{
    /// <summary>
    /// 混合对齐方式
    /// </summary>
    msoVerticalAnchorMixed = -2,
    /// <summary>
    /// 顶部对齐
    /// </summary>
    msoAnchorTop = 1,
    /// <summary>
    /// 顶部基线对齐
    /// </summary>
    msoAnchorTopBaseline = 2,
    /// <summary>
    /// 居中对齐
    /// </summary>
    msoAnchorMiddle = 3,
    /// <summary>
    /// 底部对齐
    /// </summary>
    msoAnchorBottom = 4,
    /// <summary>
    /// 底部基线对齐
    /// </summary>
    msoAnchorBottomBaseLine = 5
}