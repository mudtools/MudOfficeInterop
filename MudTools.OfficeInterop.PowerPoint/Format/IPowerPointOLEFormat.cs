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
/// 表示 PowerPoint 中 OLE 对象的格式设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointOLEFormat : IDisposable
{
    /// <summary>
    /// 获取创建此 OLE 格式设置的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此 OLE 格式设置的父对象。
    /// </summary>
    /// <value>表示此 OLE 格式设置父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取 OLE 对象支持的动作动词集合。
    /// </summary>
    /// <value>表示 OLE 动作动词的 <see cref="IPowerPointObjectVerbs"/> 集合。</value>
    IPowerPointObjectVerbs? ObjectVerbs { get; }

    /// <summary>
    /// 获取 OLE 对象的底层对象。
    /// </summary>
    /// <value>表示 OLE 对象的 <see cref="object"/>。</value>
    object? Object { get; }

    /// <summary>
    /// 获取 OLE 对象的程序标识符。
    /// </summary>
    /// <value>表示 OLE 对象 ProgID 的字符串。</value>
    string? ProgID { get; }

    /// <summary>
    /// 获取或设置 OLE 对象跟随颜色的方式。
    /// </summary>
    /// <value>表示颜色跟随方式的 <see cref="PpFollowColors"/> 枚举值。</value>
    PpFollowColors FollowColors { get; set; }

    /// <summary>
    /// 对 OLE 对象执行指定的动作动词。
    /// </summary>
    /// <param name="index">要执行的动作动词的索引。默认为 0，表示默认动作。</param>
    void DoVerb(int index = 0);

    /// <summary>
    /// 激活 OLE 对象以进行编辑。
    /// </summary>
    void Activate();
}