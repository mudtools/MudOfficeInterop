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
/// 表示 PowerPoint 幻灯片中的时间线对象，用于管理动画序列。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointTimeLine : IOfficeObject<IPowerPointTimeLine, MsPowerPoint.TimeLine>, IDisposable
{
    /// <summary>
    /// 获取创建此时间线的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="Application"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此时间线的父对象。
    /// </summary>
    /// <value>表示此时间线父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取主要动画序列。
    /// </summary>
    /// <value>表示主要动画序列的 <see cref="IPowerPointSequence"/> 对象。</value>
    IPowerPointSequence? MainSequence { get; }

    /// <summary>
    /// 获取交互式动画序列集合。
    /// </summary>
    /// <value>表示交互式动画序列集合的 <see cref="IPowerPointSequences"/> 对象。</value>
    IPowerPointSequences? InteractiveSequences { get; }
}