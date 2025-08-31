namespace MudTools.OfficeInterop;

/// <summary>
/// 指定文本框或形状中文本的水平对齐方式
/// </summary>
public enum MsoHorizontalAnchor
{
    /// <summary>
    /// 混合水平对齐模式（通常用于表示未设置或混合状态）
    /// </summary>
    msoHorizontalAnchorMixed = -2,
    
    /// <summary>
    /// 无特定水平对齐（使用默认对齐方式）
    /// </summary>
    msoAnchorNone = 1,
    
    /// <summary>
    /// 水平居中对齐
    /// </summary>
    msoAnchorCenter = 2
}