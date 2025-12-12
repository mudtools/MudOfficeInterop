//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Excel 中对象的倒影（Reflection）格式设置的封装接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.ReflectionFormat
/// 用于控制倒影的透明度、大小、模糊度、偏移等视觉效果。
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
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
    float Transparency { get; set; }

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
}