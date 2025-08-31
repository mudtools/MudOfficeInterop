namespace MudTools.OfficeInteropShapes;

/// <summary>
/// 指定线条的样式类型
/// </summary>
public enum MsoLineStyle
{
    /// <summary>
    /// 混合线条样式（用于组合图形）
    /// </summary>
    msoLineStyleMixed = -2,
    /// <summary>
    /// 单线样式
    /// </summary>
    msoLineSingle = 1,
    /// <summary>
    /// 细-细线样式（两条细线）
    /// </summary>
    msoLineThinThin = 2,
    /// <summary>
    /// 细-粗线样式（一条细线和一条粗线）
    /// </summary>
    msoLineThinThick = 3,
    /// <summary>
    /// 粗-细线样式（一条粗线和一条细线）
    /// </summary>
    msoLineThickThin = 4,
    /// <summary>
    /// 粗线夹在两条细线之间样式
    /// </summary>
    msoLineThickBetweenThin = 5
}