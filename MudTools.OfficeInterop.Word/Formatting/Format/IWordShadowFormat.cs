//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

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

    bool RotateWithShape { get; set; }

    MsoShadowStyle Style { get; set; }

    float OffsetX { get; set; }

    float OffsetY { get; set; }

    void SetOffset(float offsetX, float offsetY);

    void Clear();

    void CopyTo(IWordShadowFormat targetShadow);

    void Reset();

    void ApplyOuterShadow(float offsetX, float offsetY, float blur, int color, float transparency);

    void ApplyInnerShadow(float offsetX, float offsetY, float blur, int color, float transparency);
}