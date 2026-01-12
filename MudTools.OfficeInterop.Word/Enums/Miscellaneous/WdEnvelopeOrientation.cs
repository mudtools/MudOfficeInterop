namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定信封的方向和位置设置
/// </summary>
public enum WdEnvelopeOrientation
{
    /// <summary>
    /// 左侧纵向排列
    /// </summary>
    wdLeftPortrait,

    /// <summary>
    /// 居中纵向排列
    /// </summary>
    wdCenterPortrait,

    /// <summary>
    /// 右侧纵向排列
    /// </summary>
    wdRightPortrait,

    /// <summary>
    /// 左侧横向排列
    /// </summary>
    wdLeftLandscape,

    /// <summary>
    /// 居中横向排列
    /// </summary>
    wdCenterLandscape,

    /// <summary>
    /// 右侧横向排列
    /// </summary>
    wdRightLandscape,

    /// <summary>
    /// 左侧顺时针排列
    /// </summary>
    wdLeftClockwise,

    /// <summary>
    /// 居中顺时针排列
    /// </summary>
    wdCenterClockwise,

    /// <summary>
    /// 右侧顺时针排列
    /// </summary>
    wdRightClockwise
}