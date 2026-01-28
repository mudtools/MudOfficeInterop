//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;
/// <summary>
/// 表示 Office 中阴影格式的接口封装。
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeShadowFormat : IOfficeObject<IOfficeShadowFormat, MsCore.ShadowFormat>, IDisposable
{
    /// <summary>
    /// 获取阴影的颜色格式。
    /// </summary>
    IOfficeColorFormat? ForeColor { get; }

    /// <summary>
    /// 获取或设置阴影的透明度（0.0 到 1.0 之间）。
    /// </summary>
    float Transparency { get; set; }

    /// <summary>
    /// 获取或设置阴影的模糊度。
    /// </summary>
    float Blur { get; set; }

    /// <summary>
    /// 获取或设置阴影的水平偏移量。
    /// </summary>
    float OffsetX { get; set; }

    /// <summary>
    /// 获取或设置阴影的垂直偏移量。
    /// </summary>
    float OffsetY { get; set; }

    /// <summary>
    /// 获取或设置阴影的大小。
    /// </summary>
    float Size { get; set; }

    /// <summary>
    /// 获取或设置阴影的可见性。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置阴影是否随形状旋转。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool RotateWithShape { get; set; }

    /// <summary>
    /// 获取或设置阴影是否被遮挡。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Obscured { get; set; }

    /// <summary>
    /// 增加阴影的水平偏移量。
    /// </summary>
    /// <param name="offsetX">要增加的水平偏移量值。</param>
    void IncrementOffsetX(float offsetX);

    /// <summary>
    /// 增加阴影的垂直偏移量。
    /// </summary>
    /// <param name="offsetY">要增加的垂直偏移量值。</param>
    void IncrementOffsetY(float offsetY);

    /// <summary>
    /// 获取或设置阴影的样式。
    /// </summary>
    MsoShadowStyle Style { get; set; }
}