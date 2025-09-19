namespace MudTools.OfficeInterop;

/// <summary>
/// 指定标注线的起始点位置类型
/// </summary>
public enum MsoCalloutDropType
{
    /// <summary>
    /// 混合标注线类型
    /// </summary>
    msoCalloutDropMixed = -2,
    
    /// <summary>
    /// 自定义标注线起始点位置
    /// </summary>
    msoCalloutDropCustom = 1,
    
    /// <summary>
    /// 顶部对齐标注线起始点
    /// </summary>
    msoCalloutDropTop = 2,
    
    /// <summary>
    /// 居中对齐标注线起始点
    /// </summary>
    msoCalloutDropCenter = 3,
    
    /// <summary>
    /// 底部对齐标注线起始点
    /// </summary>
    msoCalloutDropBottom = 4
}