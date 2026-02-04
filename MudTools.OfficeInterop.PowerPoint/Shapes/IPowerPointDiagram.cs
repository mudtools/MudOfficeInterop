//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示一个图示对象，提供对图示属性和操作的访问。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointDiagram : IOfficeObject<IPowerPointDiagram, MsPowerPoint.Diagram>, IDisposable
{
    /// <summary>
    /// 获取创建此图示的应用程序。
    /// </summary>
    /// <value>表示应用程序的对象。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取创建此对象的应用程序的创建者代码。
    /// </summary>
    /// <value>创建者标识符。</value>
    int Creator { get; }

    /// <summary>
    /// 获取图示的父对象。
    /// </summary>
    /// <value>父对象，通常是形状或幻灯片。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取图示中所有节点的集合。
    /// </summary>
    /// <value>图示节点的集合。</value>
    IPowerPointDiagramNodes? Nodes { get; }

    /// <summary>
    /// 获取图示的类型。
    /// </summary>
    /// <value>图示的类型标识。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoDiagramType Type { get; }

    /// <summary>
    /// 获取或设置一个值，指示图示是否自动布局。
    /// </summary>
    /// <value>如果图示自动布局，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool AutoLayout { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示图示是否反向显示。
    /// </summary>
    /// <value>如果图示反向显示，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Reverse { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示图示是否自动应用格式。
    /// </summary>
    /// <value>如果图示自动应用格式，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool AutoFormat { get; set; }

    /// <summary>
    /// 将图示转换为指定的类型。
    /// </summary>
    /// <param name="type">要将图示转换成的目标类型。</param>
    void Convert([ComNamespace("MsCore")] MsoDiagramType type);

    /// <summary>
    /// 调整图示中文本的大小以适应节点。
    /// </summary>
    void FitText();
}