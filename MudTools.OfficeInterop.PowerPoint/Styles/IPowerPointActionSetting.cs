//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

using System;
using System.Runtime.InteropServices;

/// <summary>
/// 表示 PowerPoint 中形状的动作设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointActionSetting : IDisposable
{
    /// <summary>
    /// 获取创建此动作设置的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="Application"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此动作设置的父对象。
    /// </summary>
    /// <value>表示此动作设置父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置动作类型。
    /// </summary>
    /// <value>表示动作类型的 <see cref="PpActionType"/> 枚举值。</value>
    PpActionType Action { get; set; }

    /// <summary>
    /// 获取或设置动作动词。
    /// </summary>
    /// <value>表示动作动词的字符串。</value>
    string? ActionVerb { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否在触发动作时应用动画效果。
    /// </summary>
    /// <value>指示是否应用动画效果的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool AnimateAction { get; set; }

    /// <summary>
    /// 获取或设置要运行的程序或宏的名称。
    /// </summary>
    /// <value>表示要运行的程序或宏名称的字符串。</value>
    string? Run { get; set; }

    /// <summary>
    /// 获取或设置幻灯片放映的名称。
    /// </summary>
    /// <value>表示幻灯片放映名称的字符串。</value>
    string? SlideShowName { get; set; }

    /// <summary>
    /// 获取与此动作设置关联的超链接。
    /// </summary>
    /// <value>表示超链接的 <see cref="IPowerPointHyperlink"/> 对象。</value>
    IPowerPointHyperlink? Hyperlink { get; }

    /// <summary>
    /// 获取与此动作设置关联的声音效果。
    /// </summary>
    /// <value>表示声音效果的 <see cref="IPowerPointSoundEffect"/> 对象。</value>
    IPowerPointSoundEffect? SoundEffect { get; }

    /// <summary>
    /// 获取或设置一个值，指示在展示其他幻灯片后是否返回到原始幻灯片。
    /// </summary>
    /// <value>指示是否返回的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool ShowAndReturn { get; set; }
}