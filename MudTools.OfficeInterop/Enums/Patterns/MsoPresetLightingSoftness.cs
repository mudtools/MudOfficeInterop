namespace MudTools.OfficeInterop;

/// <summary>
/// 指定形状的预设照明柔和度
/// </summary>
public enum MsoPresetLightingSoftness
{
    /// <summary>
    /// 混合照明柔和度
    /// </summary>
    msoPresetLightingSoftnessMixed = -2,
    /// <summary>
    /// 暗淡照明
    /// </summary>
    msoLightingDim = 1,
    /// <summary>
    /// 正常照明
    /// </summary>
    msoLightingNormal = 2,
    /// <summary>
    /// 明亮照明
    /// </summary>
    msoLightingBright = 3
}