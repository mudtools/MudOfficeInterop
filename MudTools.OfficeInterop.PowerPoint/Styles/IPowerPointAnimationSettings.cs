//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 形状的动画设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointAnimationSettings : IOfficeObject<IPowerPointAnimationSettings, MsPowerPoint.AnimationSettings>, IDisposable
{
    /// <summary>
    /// 获取创建此动画设置的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此动画设置的父对象。
    /// </summary>
    /// <value>表示此动画设置父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取动画后变暗的颜色设置。
    /// </summary>
    /// <value>表示变暗颜色的 <see cref="IPowerPointColorFormat"/> 对象。</value>
    IPowerPointColorFormat? DimColor { get; }

    /// <summary>
    /// 获取动画的声音效果设置。
    /// </summary>
    /// <value>表示声音效果的 <see cref="IPowerPointSoundEffect"/> 对象。</value>
    IPowerPointSoundEffect? SoundEffect { get; }

    /// <summary>
    /// 获取或设置进入动画效果。
    /// </summary>
    /// <value>表示进入效果的 <see cref="PpEntryEffect"/> 枚举值。</value>
    PpEntryEffect EntryEffect { get; set; }

    /// <summary>
    /// 获取或设置动画后的效果。
    /// </summary>
    /// <value>表示动画后效果的 <see cref="PpAfterEffect"/> 枚举值。</value>
    PpAfterEffect AfterEffect { get; set; }

    /// <summary>
    /// 获取或设置动画的顺序。
    /// </summary>
    /// <value>表示动画顺序的整数值。</value>
    int AnimationOrder { get; set; }

    /// <summary>
    /// 获取或设置动画的推进模式。
    /// </summary>
    /// <value>表示推进模式的 <see cref="PpAdvanceMode"/> 枚举值。</value>
    PpAdvanceMode AdvanceMode { get; set; }

    /// <summary>
    /// 获取或设置自动推进动画的时间（以秒为单位）。
    /// </summary>
    /// <value>表示自动推进时间的浮点数。</value>
    float AdvanceTime { get; set; }

    /// <summary>
    /// 获取动画的播放设置。
    /// </summary>
    /// <value>表示播放设置的 <see cref="IPowerPointPlaySettings"/> 对象。</value>
    IPowerPointPlaySettings? PlaySettings { get; }

    /// <summary>
    /// 获取或设置文本动画的层级效果。
    /// </summary>
    /// <value>表示文本层级效果的 <see cref="PpTextLevelEffect"/> 枚举值。</value>
    PpTextLevelEffect TextLevelEffect { get; set; }

    /// <summary>
    /// 获取或设置文本动画的单位效果。
    /// </summary>
    /// <value>表示文本单位效果的 <see cref="PpTextUnitEffect"/> 枚举值。</value>
    PpTextUnitEffect TextUnitEffect { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否启用动画。
    /// </summary>
    /// <value>指示是否启用动画的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Animate { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否启用背景动画。
    /// </summary>
    /// <value>指示是否启用背景动画的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool AnimateBackground { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否以相反顺序动画文本。
    /// </summary>
    /// <value>指示是否以相反顺序动画文本的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool AnimateTextInReverse { get; set; }

    /// <summary>
    /// 获取或设置图表动画的单位效果。
    /// </summary>
    /// <value>表示图表单位效果的 <see cref="PpChartUnitEffect"/> 枚举值。</value>
    PpChartUnitEffect ChartUnitEffect { get; set; }
}
