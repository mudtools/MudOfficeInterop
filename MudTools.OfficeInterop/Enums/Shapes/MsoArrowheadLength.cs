namespace MudTools.OfficeInterop;

/// <summary>
/// 指定线条端点箭头长度的枚举，用于定义形状线条端点的箭头长度
/// </summary>
public enum MsoArrowheadLength
{
    /// <summary>
    /// 混合箭头长度（用于表示多种箭头长度的组合）
    /// </summary>
    msoArrowheadLengthMixed = -2,
    
    /// <summary>
    /// 短箭头长度
    /// </summary>
    msoArrowheadShort = 1,
    
    /// <summary>
    /// 中等箭头长度
    /// </summary>
    msoArrowheadLengthMedium = 2,
    
    /// <summary>
    /// 长箭头长度
    /// </summary>
    msoArrowheadLong = 3
}