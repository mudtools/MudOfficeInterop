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
/// 表示 PowerPoint 幻灯片或备注页中的页眉和页脚集合。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointHeadersFooters : IOfficeObject<IPowerPointHeadersFooters, MsPowerPoint.HeadersFooters>, IDisposable
{
    /// <summary>
    /// 获取创建此页眉页脚集合的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此页眉页脚集合的父对象。
    /// </summary>
    /// <value>表示此页眉页脚集合父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取日期和时间页脚对象。
    /// </summary>
    /// <value>表示日期和时间设置的 <see cref="IPowerPointHeaderFooter"/> 对象。</value>
    IPowerPointHeaderFooter? DateAndTime { get; }

    /// <summary>
    /// 获取幻灯片编号页脚对象。
    /// </summary>
    /// <value>表示幻灯片编号设置的 <see cref="IPowerPointHeaderFooter"/> 对象。</value>
    IPowerPointHeaderFooter? SlideNumber { get; }

    /// <summary>
    /// 获取页眉对象。
    /// </summary>
    /// <value>表示页眉设置的 <see cref="IPowerPointHeaderFooter"/> 对象。</value>
    IPowerPointHeaderFooter? Header { get; }

    /// <summary>
    /// 获取页脚对象。
    /// </summary>
    /// <value>表示页脚设置的 <see cref="IPowerPointHeaderFooter"/> 对象。</value>
    IPowerPointHeaderFooter? Footer { get; }

    /// <summary>
    /// 获取或设置一个值，指示是否在标题幻灯片上显示页眉和页脚。
    /// </summary>
    /// <value>指示是否在标题幻灯片上显示的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool DisplayOnTitleSlide { get; set; }

    /// <summary>
    /// 清除所有页眉和页脚设置。
    /// </summary>
    void Clear();
}