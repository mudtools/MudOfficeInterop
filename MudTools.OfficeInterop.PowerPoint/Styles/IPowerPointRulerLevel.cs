//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 标尺中的层级设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointRulerLevel : IOfficeObject<IPowerPointRulerLevel, MsPowerPoint.RulerLevel>, IDisposable
{
    /// <summary>
    /// 获取创建此标尺层级的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此标尺层级的父对象。
    /// </summary>
    /// <value>表示此标尺层级父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置首行缩进（以磅为单位）。
    /// </summary>
    /// <value>表示首行缩进的浮点数。</value>
    float FirstMargin { get; set; }

    /// <summary>
    /// 获取或设置左缩进（以磅为单位）。
    /// </summary>
    /// <value>表示左缩进的浮点数。</value>
    float LeftMargin { get; set; }
}