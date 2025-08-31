namespace MudTools.OfficeInterop;

/// <summary>
/// 指定形状线条端点箭头头部的宽度
/// </summary>
public enum MsoArrowheadWidth
{
    /// <summary>
    /// 混合箭头头部宽度（用于包含多种宽度的组合）
    /// </summary>
    msoArrowheadWidthMixed = -2,
    /// <summary>
    /// 窄箭头头部宽度
    /// </summary>
    msoArrowheadNarrow = 1,
    /// <summary>
    /// 中等箭头头部宽度
    /// </summary>
    msoArrowheadWidthMedium = 2,
    /// <summary>
    /// 宽箭头头部宽度
    /// </summary>
    msoArrowheadWide = 3
}