//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 形状的阴影格式设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointShadowFormat : IOfficeObject<IPowerPointShadowFormat, MsPowerPoint.ShadowFormat>, IDisposable
{
    /// <summary>
    /// 获取创建此阴影格式设置的应用程序实例。
    /// </summary>
    /// <value>表示应用程序的 <see cref="IPowerPointApplication"/>。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取创建此阴影格式设置的应用程序的创建者代码。
    /// </summary>
    /// <value>表示创建者代码的整数值。</value>
    int Creator { get; }

    /// <summary>
    /// 获取此阴影格式设置的父对象。
    /// </summary>
    /// <value>表示此阴影格式设置父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 按指定增量增加阴影的水平偏移量。
    /// </summary>
    /// <param name="increment">水平偏移量增量（以磅为单位）。</param>
    void IncrementOffsetX(float increment);

    /// <summary>
    /// 按指定增量增加阴影的垂直偏移量。
    /// </summary>
    /// <param name="increment">垂直偏移量增量（以磅为单位）。</param>
    void IncrementOffsetY(float increment);

    /// <summary>
    /// 获取或设置阴影的前景颜色。
    /// </summary>
    /// <value>表示阴影颜色的 <see cref="IPowerPointColorFormat"/> 对象。</value>
    IPowerPointColorFormat? ForeColor { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示阴影是否被形状遮挡。
    /// </summary>
    /// <value>指示阴影是否被遮挡的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Obscured { get; set; }

    /// <summary>
    /// 获取或设置阴影的水平偏移量（以磅为单位）。
    /// </summary>
    /// <value>表示水平偏移量的浮点数。</value>
    float OffsetX { get; set; }

    /// <summary>
    /// 获取或设置阴影的垂直偏移量（以磅为单位）。
    /// </summary>
    /// <value>表示垂直偏移量的浮点数。</value>
    float OffsetY { get; set; }

    /// <summary>
    /// 获取或设置阴影的透明度。
    /// </summary>
    /// <value>表示透明度值的浮点数（0.0 到 1.0）。</value>
    float Transparency { get; set; }

    /// <summary>
    /// 获取或设置阴影的类型。
    /// </summary>
    /// <value>表示阴影类型的 <see cref="MsoShadowType"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoShadowType Type { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示阴影是否可见。
    /// </summary>
    /// <value>指示阴影是否可见的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置阴影的样式。
    /// </summary>
    /// <value>表示阴影样式的 <see cref="MsoShadowStyle"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoShadowStyle Style { get; set; }

    /// <summary>
    /// 获取或设置阴影的模糊程度。
    /// </summary>
    /// <value>表示模糊程度的浮点数（以磅为单位）。</value>
    float Blur { get; set; }

    /// <summary>
    /// 获取或设置阴影的大小。
    /// </summary>
    /// <value>表示阴影大小的浮点数（以磅为单位）。</value>
    float Size { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示阴影是否随形状旋转。
    /// </summary>
    /// <value>指示阴影是否随形状旋转的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool RotateWithShape { get; set; }
}