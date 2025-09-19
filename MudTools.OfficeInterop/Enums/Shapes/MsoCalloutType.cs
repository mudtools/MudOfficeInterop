namespace MudTools.OfficeInterop;

/// <summary>
/// 指定标注线的类型，用于连接文本框和图形对象
/// </summary>
public enum MsoCalloutType
{
    /// <summary>
    /// 混合标注线类型
    /// </summary>
    msoCalloutMixed = -2,
    
    /// <summary>
    /// 第一种标注线类型
    /// </summary>
    msoCalloutOne = 1,
    
    /// <summary>
    /// 第二种标注线类型
    /// </summary>
    msoCalloutTwo = 2,
    
    /// <summary>
    /// 第三种标注线类型
    /// </summary>
    msoCalloutThree = 3,
    
    /// <summary>
    /// 第四种标注线类型
    /// </summary>
    msoCalloutFour = 4
}