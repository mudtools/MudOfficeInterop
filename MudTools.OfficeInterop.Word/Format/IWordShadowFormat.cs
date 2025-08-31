namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Core.ShadowFormat 的接口，用于操作阴影格式。
/// </summary>
public interface IWordShadowFormat : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取阴影的前景颜色格式。
    /// </summary>
    IWordColorFormat ForeColor { get; }

    /// <summary>
    /// 获取或设置阴影的透明度（0.0到1.0之间）。
    /// </summary>
    float Transparency { get; set; }

    /// <summary>
    /// 获取或设置阴影是否可见。
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置阴影的类型。
    /// </summary>
    MsoShadowType Type { get; set; }

    /// <summary>
    /// 获取或设置阴影的模糊度。
    /// </summary>
    float Blur { get; set; }

    /// <summary>
    /// 获取或设置阴影的大小。
    /// </summary>
    float Size { get; set; }

    /// <summary>
    /// 获取或设置阴影的旋转角度。
    /// </summary>
    float RotateWithShape { get; set; }

    /// <summary>
    /// 获取或设置阴影的距离。
    /// </summary>
    float Distance { get; set; }

    /// <summary>
    /// 获取或设置阴影的角度。
    /// </summary>
    float Angle { get; set; }
}