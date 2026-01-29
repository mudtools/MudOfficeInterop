//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 动画效果的时间设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointTiming : IDisposable
{
    /// <summary>
    /// 获取创建此时间设置的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此时间设置的父对象。
    /// </summary>
    /// <value>表示此时间设置父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置动画效果的持续时间（以秒为单位）。
    /// </summary>
    /// <value>表示持续时间的浮点数。</value>
    float Duration { get; set; }

    /// <summary>
    /// 获取或设置动画效果的触发类型。
    /// </summary>
    /// <value>表示触发类型的 <see cref="MsoAnimTriggerType"/> 枚举值。</value>
    MsoAnimTriggerType TriggerType { get; set; }

    /// <summary>
    /// 获取或设置动画效果的触发延迟时间（以秒为单位）。
    /// </summary>
    /// <value>表示触发延迟时间的浮点数。</value>
    float TriggerDelayTime { get; set; }

    /// <summary>
    /// 获取或设置触发动画效果的形状。
    /// </summary>
    /// <value>表示触发形状的 <see cref="IPowerPointShape"/> 对象。</value>
    IPowerPointShape? TriggerShape { get; set; }

    /// <summary>
    /// 获取或设置动画效果的重复次数。
    /// </summary>
    /// <value>表示重复次数的整数值。</value>
    int RepeatCount { get; set; }

    /// <summary>
    /// 获取或设置动画效果的重复持续时间（以秒为单位）。
    /// </summary>
    /// <value>表示重复持续时间的浮点数。</value>
    float RepeatDuration { get; set; }

    /// <summary>
    /// 获取或设置动画效果的速度。
    /// </summary>
    /// <value>表示速度的浮点数。</value>
    float Speed { get; set; }

    /// <summary>
    /// 获取或设置动画效果的加速值（0.0 到 1.0）。
    /// </summary>
    /// <value>表示加速值的浮点数。</value>
    float Accelerate { get; set; }

    /// <summary>
    /// 获取或设置动画效果的减速值（0.0 到 1.0）。
    /// </summary>
    /// <value>表示减速值的浮点数。</value>
    float Decelerate { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示动画效果是否自动反向播放。
    /// </summary>
    /// <value>指示是否自动反向播放的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool AutoReverse { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示动画效果是否平滑开始。
    /// </summary>
    /// <value>指示是否平滑开始的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool SmoothStart { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示动画效果是否平滑结束。
    /// </summary>
    /// <value>指示是否平滑结束的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool SmoothEnd { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示动画效果是否在结束时倒回。
    /// </summary>
    /// <value>指示是否在结束时倒回的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool RewindAtEnd { get; set; }

    /// <summary>
    /// 获取或设置动画效果的重启方式。
    /// </summary>
    /// <value>表示重启方式的 <see cref="MsoAnimEffectRestart"/> 枚举值。</value>
    MsoAnimEffectRestart Restart { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示动画效果是否在结束时弹跳。
    /// </summary>
    /// <value>指示是否在结束时弹跳的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool BounceEnd { get; set; }

    /// <summary>
    /// 获取或设置动画效果结束时的弹跳强度。
    /// </summary>
    /// <value>表示弹跳强度的浮点数。</value>
    float BounceEndIntensity { get; set; }

    /// <summary>
    /// 获取或设置触发动画效果的书签名称。
    /// </summary>
    /// <value>表示书签名称的字符串。</value>
    string? TriggerBookmark { get; set; }
}