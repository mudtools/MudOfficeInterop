namespace MudTools.OfficeInterop;

/// <summary>
/// 指定标注线的角度类型
/// </summary>
public enum MsoCalloutAngleType
{
    /// <summary>
    /// 混合角度类型
    /// </summary>
    msoCalloutAngleMixed = -2,
    /// <summary>
    /// 自动角度类型
    /// </summary>
    msoCalloutAngleAutomatic = 1,
    /// <summary>
    /// 30度角
    /// </summary>
    msoCalloutAngle30 = 2,
    /// <summary>
    /// 45度角
    /// </summary>
    msoCalloutAngle45 = 3,
    /// <summary>
    /// 60度角
    /// </summary>
    msoCalloutAngle60 = 4,
    /// <summary>
    /// 90度角
    /// </summary>
    msoCalloutAngle90 = 5
}