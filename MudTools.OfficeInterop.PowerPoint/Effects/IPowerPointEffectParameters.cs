//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 动画效果的参数设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointEffectParameters : IDisposable
{
    /// <summary>
    /// 获取创建此效果参数的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此效果参数的父对象。
    /// </summary>
    /// <value>表示此效果参数父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置动画效果的方向。
    /// </summary>
    /// <value>表示动画方向的 <see cref="MsoAnimDirection"/> 枚举值。</value>
    MsoAnimDirection Direction { get; set; }

    /// <summary>
    /// 获取或设置动画效果的幅度。
    /// </summary>
    /// <value>表示幅度的浮点数。</value>
    float Amount { get; set; }

    /// <summary>
    /// 获取或设置动画效果的大小。
    /// </summary>
    /// <value>表示大小的浮点数。</value>
    float Size { get; set; }

    /// <summary>
    /// 获取动画效果的第二种颜色设置。
    /// </summary>
    /// <value>表示第二种颜色的 <see cref="IPowerPointColorFormat"/> 对象。</value>
    IPowerPointColorFormat? Color2 { get; }

    /// <summary>
    /// 获取或设置一个值，指示动画效果参数是否为相对值。
    /// </summary>
    /// <value>指示是否为相对值的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Relative { get; set; }

    /// <summary>
    /// 获取或设置动画效果的字体名称。
    /// </summary>
    /// <value>表示字体名称的字符串。</value>
    string? FontName { get; set; }
}