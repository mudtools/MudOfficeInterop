namespace MudTools.OfficeInterop;

/// <summary>
/// 指定在缩放操作期间的参考点位置
/// </summary>
public enum MsoScaleFrom
{
    /// <summary>
    /// 从左上角作为参考点进行缩放
    /// </summary>
    msoScaleFromTopLeft,

    /// <summary>
    /// 从中心位置作为参考点进行缩放
    /// </summary>
    msoScaleFromMiddle,

    /// <summary>
    /// 从右下角作为参考点进行缩放
    /// </summary>
    msoScaleFromBottomRight
}