
namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Excel 中对象的倒影（Reflection）格式设置的封装接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.ReflectionFormat
/// 用于控制倒影的透明度、大小、模糊度、偏移等视觉效果。
/// </summary>
public interface IOfficeReflectionFormat : IDisposable
{
    /// <summary>
    /// 获取或设置倒影的类型（无倒影、预设倒影样式等）。
    /// 使用 <see cref="MsoReflectionType"/> 枚举。
    /// </summary>
    MsoReflectionType Type { get; set; }

    /// <summary>
    /// 获取或设置倒影的透明度（0-100，0=完全不透明，100=完全透明）。
    /// 内部自动转换为 COM 所需的 0.0~1.0 浮点值。
    /// </summary>
    int Transparency { get; set; }

    /// <summary>
    /// 获取或设置倒影的大小比例（0.0~1.0，1.0=100% 原图高度）。
    /// 例如：0.5 表示倒影高度为原对象的一半。
    /// </summary>
    float Size { get; set; }

    /// <summary>
    /// 获取或设置倒影的模糊程度（单位：磅）。
    /// 值越大，倒影边缘越模糊。
    /// </summary>
    float Blur { get; set; }

    /// <summary>
    /// 获取或设置倒影与原对象的垂直距离（单位：磅）。
    /// 正值表示倒影在对象下方。
    /// </summary>
    float Offset { get; set; }

    /// <summary>
    /// 获取倒影效果是否已启用（Type != msoReflectionTypeNone）。
    /// </summary>
    bool Visible { get; }
}