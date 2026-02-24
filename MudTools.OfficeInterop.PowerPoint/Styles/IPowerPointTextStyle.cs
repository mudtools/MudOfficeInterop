//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;


/// <summary>
/// 表示 PowerPoint 幻灯片母版中的文本样式。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointTextStyle : IOfficeObject<IPowerPointTextStyle, MsPowerPoint.TextStyle>, IDisposable
{
    /// <summary>
    /// 获取创建此文本样式的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此文本样式的父对象。
    /// </summary>
    /// <value>表示此文本样式父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取此文本样式的标尺对象。
    /// </summary>
    /// <value>表示标尺的 <see cref="IPowerPointRuler"/> 对象。</value>
    IPowerPointRuler? Ruler { get; }

    /// <summary>
    /// 获取此文本样式的文本框对象。
    /// </summary>
    /// <value>表示文本框的 <see cref="IPowerPointTextFrame"/> 对象。</value>
    IPowerPointTextFrame? TextFrame { get; }

    /// <summary>
    /// 获取此文本样式的层级集合。
    /// </summary>
    /// <value>表示文本样式层级的 <see cref="IPowerPointTextStyleLevels"/> 对象。</value>
    IPowerPointTextStyleLevels? Levels { get; }
}