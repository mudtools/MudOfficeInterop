//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 动画效果的信息。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointEffectInformation : IDisposable
{
    /// <summary>
    /// 获取创建此效果信息的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此效果信息的父对象。
    /// </summary>
    /// <value>表示此效果信息父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取动画效果的后续效果类型。
    /// </summary>
    /// <value>表示后续效果的 <see cref="MsoAnimAfterEffect"/> 枚举值。</value>
    MsoAnimAfterEffect AfterEffect { get; }

    /// <summary>
    /// 获取一个值，指示是否启用背景动画。
    /// </summary>
    /// <value>指示是否启用背景动画的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool AnimateBackground { get; }

    /// <summary>
    /// 获取一个值，指示是否以相反顺序动画文本。
    /// </summary>
    /// <value>指示是否以相反顺序动画文本的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool AnimateTextInReverse { get; }

    /// <summary>
    /// 获取按层级构建动画的效果类型。
    /// </summary>
    /// <value>表示按层级动画效果的 <see cref="MsoAnimateByLevel"/> 枚举值。</value>
    MsoAnimateByLevel BuildByLevelEffect { get; }

    /// <summary>
    /// 获取动画后变暗的颜色设置。
    /// </summary>
    /// <value>表示变暗颜色的 <see cref="IPowerPointColorFormat"/> 对象。</value>
    IPowerPointColorFormat? Dim { get; }

    /// <summary>
    /// 获取动画的播放设置。
    /// </summary>
    /// <value>表示播放设置的 <see cref="IPowerPointPlaySettings"/> 对象。</value>
    IPowerPointPlaySettings? PlaySettings { get; }

    /// <summary>
    /// 获取动画的声音效果设置。
    /// </summary>
    /// <value>表示声音效果的 <see cref="IPowerPointSoundEffect"/> 对象。</value>
    IPowerPointSoundEffect? SoundEffect { get; }

    /// <summary>
    /// 获取动画效果的文本单位效果类型。
    /// </summary>
    /// <value>表示文本单位效果的 <see cref="MsoAnimTextUnitEffect"/> 枚举值。</value>
    MsoAnimTextUnitEffect TextUnitEffect { get; }
}