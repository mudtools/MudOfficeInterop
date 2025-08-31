namespace MudTools.OfficeInterop;

/// <summary>
/// 指定Office应用程序中形状的阴影样式
/// </summary>
public enum MsoShadowStyle
{
    /// <summary>
    /// 混合阴影样式（通常用于表示多种样式的组合）
    /// </summary>
    msoShadowStyleMixed = -2,

    /// <summary>
    /// 内阴影样式，阴影显示在形状内部
    /// </summary>
    msoShadowStyleInnerShadow = 1,

    /// <summary>
    /// 外阴影样式，阴影显示在形状外部
    /// </summary>
    msoShadowStyleOuterShadow = 2
}