//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 动画中的滤镜效果设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointFilterEffect : IOfficeObject<IPowerPointFilterEffect, MsPowerPoint.FilterEffect>, IDisposable
{
    /// <summary>
    /// 获取创建此滤镜效果的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此滤镜效果的父对象。
    /// </summary>
    /// <value>表示此滤镜效果父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置滤镜效果的类型。
    /// </summary>
    /// <value>表示滤镜类型的 <see cref="MsoAnimFilterEffectType"/> 枚举值。</value>
    MsoAnimFilterEffectType Type { get; set; }

    /// <summary>
    /// 获取或设置滤镜效果的子类型。
    /// </summary>
    /// <value>表示滤镜子类型的 <see cref="MsoAnimFilterEffectSubtype"/> 枚举值。</value>
    MsoAnimFilterEffectSubtype Subtype { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示滤镜效果是否揭示内容。
    /// </summary>
    /// <value>指示是否揭示内容的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Reveal { get; set; }
}