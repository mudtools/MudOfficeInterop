namespace MudTools.OfficeInterop;

/// <summary>
/// 指定三维挤出效果的颜色类型
/// </summary>
public enum MsoExtrusionColorType
{
    /// <summary>
    /// 混合颜色类型
    /// </summary>
    msoExtrusionColorTypeMixed = -2,
    /// <summary>
    /// 自动颜色（使用对象的自动颜色）
    /// </summary>
    msoExtrusionColorAutomatic = 1,
    /// <summary>
    /// 自定义颜色（使用指定的自定义颜色）
    /// </summary>
    msoExtrusionColorCustom = 2
}