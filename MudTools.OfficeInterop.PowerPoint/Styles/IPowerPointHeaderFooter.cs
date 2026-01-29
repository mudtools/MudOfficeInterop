//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 幻灯片或备注页中的页眉或页脚。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointHeaderFooter : IDisposable
{
    /// <summary>
    /// 获取创建此页眉页脚的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此页眉页脚的父对象。
    /// </summary>
    /// <value>表示此页眉页脚父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置一个值，指示页眉或页脚是否可见。
    /// </summary>
    /// <value>指示是否可见的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置页眉或页脚的文本内容。
    /// </summary>
    /// <value>表示页眉页脚文本内容的字符串。</value>
    string? Text { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否使用指定的日期/时间格式。
    /// </summary>
    /// <value>指示是否使用格式的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool UseFormat { get; set; }

    /// <summary>
    /// 获取或设置日期/时间的显示格式。
    /// </summary>
    /// <value>表示日期时间格式的 <see cref="PpDateTimeFormat"/> 枚举值。</value>
    PpDateTimeFormat Format { get; set; }
}