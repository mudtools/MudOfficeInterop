namespace MudTools.OfficeInterop;

/// <summary>
/// 指定形状的挤压效果方向
/// </summary>
public enum MsoPresetExtrusionDirection
{
    /// <summary>
    /// 混合挤压方向
    /// </summary>
    msoPresetExtrusionDirectionMixed = -2,
    /// <summary>
    /// 右下方向挤压
    /// </summary>
    msoExtrusionBottomRight = 1,
    /// <summary>
    /// 向下挤压
    /// </summary>
    msoExtrusionBottom = 2,
    /// <summary>
    /// 左下方向挤压
    /// </summary>
    msoExtrusionBottomLeft = 3,
    /// <summary>
    /// 向右挤压
    /// </summary>
    msoExtrusionRight = 4,
    /// <summary>
    /// 无挤压效果
    /// </summary>
    msoExtrusionNone = 5,
    /// <summary>
    /// 向左挤压
    /// </summary>
    msoExtrusionLeft = 6,
    /// <summary>
    /// 右上方向挤压
    /// </summary>
    msoExtrusionTopRight = 7,
    /// <summary>
    /// 向上挤压
    /// </summary>
    msoExtrusionTop = 8,
    /// <summary>
    /// 左上方向挤压
    /// </summary>
    msoExtrusionTopLeft = 9
}