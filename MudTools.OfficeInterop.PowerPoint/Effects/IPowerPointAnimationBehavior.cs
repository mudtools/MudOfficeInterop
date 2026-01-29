//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 动画效果的行为设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointAnimationBehavior : IDisposable
{
    /// <summary>
    /// 获取创建此动画行为的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="Application"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此动画行为的父对象。
    /// </summary>
    /// <value>表示此动画行为父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置动画行为的叠加方式。
    /// </summary>
    /// <value>表示叠加方式的 <see cref="MsoAnimAdditive"/> 枚举值。</value>
    MsoAnimAdditive Additive { get; set; }

    /// <summary>
    /// 获取或设置动画行为的累积方式。
    /// </summary>
    /// <value>表示累积方式的 <see cref="MsoAnimAccumulate"/> 枚举值。</value>
    MsoAnimAccumulate Accumulate { get; set; }

    /// <summary>
    /// 获取或设置动画行为的类型。
    /// </summary>
    /// <value>表示动画类型的 <see cref="MsoAnimType"/> 枚举值。</value>
    MsoAnimType Type { get; set; }

    /// <summary>
    /// 获取此动画行为的运动效果设置。
    /// </summary>
    /// <value>表示运动效果的 <see cref="IPowerPointMotionEffect"/> 对象。</value>
    IPowerPointMotionEffect? MotionEffect { get; }

    /// <summary>
    /// 获取此动画行为的颜色效果设置。
    /// </summary>
    /// <value>表示颜色效果的 <see cref="IPowerPointColorEffect"/> 对象。</value>
    IPowerPointColorEffect? ColorEffect { get; }

    /// <summary>
    /// 获取此动画行为的缩放效果设置。
    /// </summary>
    /// <value>表示缩放效果的 <see cref="IPowerPointScaleEffect"/> 对象。</value>
    IPowerPointScaleEffect? ScaleEffect { get; }

    /// <summary>
    /// 获取此动画行为的旋转效果设置。
    /// </summary>
    /// <value>表示旋转效果的 <see cref="IPowerPointRotationEffect"/> 对象。</value>
    IPowerPointRotationEffect? RotationEffect { get; }

    /// <summary>
    /// 获取此动画行为的属性效果设置。
    /// </summary>
    /// <value>表示属性效果的 <see cref="IPowerPointPropertyEffect"/> 对象。</value>
    IPowerPointPropertyEffect? PropertyEffect { get; }

    /// <summary>
    /// 获取此动画行为的时间设置。
    /// </summary>
    /// <value>表示时间设置的 <see cref="IPowerPointTiming"/> 对象。</value>
    IPowerPointTiming? Timing { get; }

    /// <summary>
    /// 删除此动画行为。
    /// </summary>
    void Delete();

    /// <summary>
    /// 获取此动画行为的命令效果设置。
    /// </summary>
    /// <value>表示命令效果的 <see cref="IPowerPointCommandEffect"/> 对象。</value>
    IPowerPointCommandEffect? CommandEffect { get; }

    /// <summary>
    /// 获取此动画行为的滤镜效果设置。
    /// </summary>
    /// <value>表示滤镜效果的 <see cref="IPowerPointFilterEffect"/> 对象。</value>
    IPowerPointFilterEffect? FilterEffect { get; }

    /// <summary>
    /// 获取此动画行为的设置效果。
    /// </summary>
    /// <value>表示设置效果的 <see cref="IPowerPointSetEffect"/> 对象。</value>
    IPowerPointSetEffect? SetEffect { get; }
}