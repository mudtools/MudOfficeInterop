//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;


/// <summary>
/// 表示 PowerPoint 中媒体对象的播放设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointPlaySettings : IDisposable
{
    /// <summary>
    /// 获取创建此播放设置的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此播放设置的父对象。
    /// </summary>
    /// <value>表示此播放设置父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置播放动作动词。
    /// </summary>
    /// <value>表示动作动词的字符串。</value>
    string? ActionVerb { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在不播放时是否隐藏媒体对象。
    /// </summary>
    /// <value>指示是否在不播放时隐藏的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool HideWhileNotPlaying { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否循环播放直到停止。
    /// </summary>
    /// <value>指示是否循环播放的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool LoopUntilStopped { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否在进入幻灯片时自动播放。
    /// </summary>
    /// <value>指示是否自动播放的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool PlayOnEntry { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示播放完成后是否倒回媒体文件。
    /// </summary>
    /// <value>指示是否倒回媒体的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool RewindMovie { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在播放媒体时是否暂停幻灯片动画。
    /// </summary>
    /// <value>指示是否暂停动画的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool PauseAnimation { get; set; }

    /// <summary>
    /// 获取或设置在播放指定数量的幻灯片后停止播放的值。
    /// </summary>
    /// <value>表示在播放多少张幻灯片后停止的整数值。</value>
    int StopAfterSlides { get; set; }
}