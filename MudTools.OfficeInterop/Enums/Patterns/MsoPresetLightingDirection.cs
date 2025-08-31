namespace MudTools.OfficeInterop;


/// <summary>
/// 指定形状的预设照明方向
/// </summary>
public enum MsoPresetLightingDirection
{
    /// <summary>
    /// 混合照明方向
    /// </summary>
    msoPresetLightingDirectionMixed = -2,
    /// <summary>
    /// 顶部左侧照明
    /// </summary>
    msoLightingTopLeft = 1,
    /// <summary>
    /// 顶部照明
    /// </summary>
    msoLightingTop = 2,
    /// <summary>
    /// 顶部右侧照明
    /// </summary>
    msoLightingTopRight = 3,
    /// <summary>
    /// 左侧照明
    /// </summary>
    msoLightingLeft = 4,
    /// <summary>
    /// 无照明
    /// </summary>
    msoLightingNone = 5,
    /// <summary>
    /// 右侧照明
    /// </summary>
    msoLightingRight = 6,
    /// <summary>
    /// 底部左侧照明
    /// </summary>
    msoLightingBottomLeft = 7,
    /// <summary>
    /// 底部照明
    /// </summary>
    msoLightingBottom = 8,
    /// <summary>
    /// 底部右侧照明
    /// </summary>
    msoLightingBottomRight = 9
}