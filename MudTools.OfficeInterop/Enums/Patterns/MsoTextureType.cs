namespace MudTools.OfficeInterop;

/// <summary>
/// 指定纹理类型，用于Office应用程序中的图案填充
/// </summary>
public enum MsoTextureType
{
    /// <summary>
    /// 混合纹理类型，表示多种纹理类型的组合
    /// </summary>
    msoTextureTypeMixed = -2,

    /// <summary>
    /// 预设纹理类型，表示系统提供的标准纹理
    /// </summary>
    msoTexturePreset = 1,

    /// <summary>
    /// 用户自定义纹理类型，表示用户提供的自定义纹理
    /// </summary>
    msoTextureUserDefined = 2
}