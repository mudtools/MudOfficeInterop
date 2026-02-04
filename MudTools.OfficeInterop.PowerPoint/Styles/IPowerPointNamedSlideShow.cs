//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 演示文稿中的命名幻灯片放映。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointNamedSlideShow : IDisposable
{
    /// <summary>
    /// 获取创建此命名幻灯片放映的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此命名幻灯片放映的父对象。
    /// </summary>
    /// <value>表示此命名幻灯片放映父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取命名幻灯片放映的名称。
    /// </summary>
    /// <value>表示命名幻灯片放映名称的字符串。</value>
    string? Name { get; }

    /// <summary>
    /// 删除此命名幻灯片放映。
    /// </summary>
    void Delete();

    /// <summary>
    /// 获取命名幻灯片放映中包含的幻灯片标识符数组。
    /// </summary>
    /// <value>表示幻灯片标识符数组的对象。</value>
    object? SlideIDs { get; }

    /// <summary>
    /// 获取命名幻灯片放映中包含的幻灯片数量。
    /// </summary>
    /// <value>表示幻灯片数量的整数值。</value>
    int Count { get; }
}