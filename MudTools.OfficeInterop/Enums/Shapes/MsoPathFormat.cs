namespace MudTools.OfficeInterop;

/// <summary>
/// 指定文件或文件夹路径的格式类型
/// </summary>
public enum MsoPathFormat
{
    /// <summary>
    /// 代表混合格式
    /// </summary>
    msoPathTypeMixed = -2,
    
    /// <summary>
    /// 代表无格式
    /// </summary>
    msoPathTypeNone = 0,
    
    /// <summary>
    /// 路径格式类型1
    /// </summary>
    msoPathType1 = 1,
    
    /// <summary>
    /// 路径格式类型2
    /// </summary>
    msoPathType2 = 2,
    
    /// <summary>
    /// 路径格式类型3
    /// </summary>
    msoPathType3 = 3,
    
    /// <summary>
    /// 路径格式类型4
    /// </summary>
    msoPathType4 = 4
}