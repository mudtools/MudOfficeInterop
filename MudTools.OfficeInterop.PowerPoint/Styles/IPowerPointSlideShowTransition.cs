//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;



/// <summary>
/// 表示 PowerPoint 幻灯片的切换效果设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointSlideShowTransition : IOfficeObject<IPowerPointSlideShowTransition, MsPowerPoint.SlideShowTransition>, IDisposable
{
    /// <summary>
    /// 获取创建此切换效果设置的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此切换效果设置的父对象。
    /// </summary>
    /// <value>表示此切换效果设置父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置一个值，指示是否在单击时推进到下一张幻灯片。
    /// </summary>
    /// <value>指示是否在单击时推进的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool AdvanceOnClick { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否在指定时间后自动推进到下一张幻灯片。
    /// </summary>
    /// <value>指示是否自动推进的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool AdvanceOnTime { get; set; }

    /// <summary>
    /// 获取或设置自动推进的时间（以秒为单位）。
    /// </summary>
    /// <value>表示自动推进时间的浮点数。</value>
    float AdvanceTime { get; set; }

    /// <summary>
    /// 获取或设置切换效果的进入效果类型。
    /// </summary>
    /// <value>表示进入效果类型的 <see cref="PpEntryEffect"/> 枚举值。</value>
    PpEntryEffect EntryEffect { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示幻灯片是否隐藏。
    /// </summary>
    /// <value>指示幻灯片是否隐藏的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Hidden { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示声音是否循环播放直到下一个幻灯片开始。
    /// </summary>
    /// <value>指示声音是否循环播放的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool LoopSoundUntilNext { get; set; }

    /// <summary>
    /// 获取切换效果的声音效果设置。
    /// </summary>
    /// <value>表示声音效果的 <see cref="IPowerPointSoundEffect"/> 对象。</value>
    IPowerPointSoundEffect? SoundEffect { get; }

    /// <summary>
    /// 获取或设置切换效果的速度。
    /// </summary>
    /// <value>表示切换速度的 <see cref="PpTransitionSpeed"/> 枚举值。</value>
    PpTransitionSpeed Speed { get; set; }

    /// <summary>
    /// 获取或设置切换效果的持续时间（以秒为单位）。
    /// </summary>
    /// <value>表示持续时间的浮点数。</value>
    float Duration { get; set; }
}
