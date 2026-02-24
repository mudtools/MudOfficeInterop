//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 中当前选中的对象。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointSelection : IOfficeObject<IPowerPointSelection, MsPowerPoint.Selection>, IDisposable
{
    /// <summary>
    /// 获取创建此选择对象的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此选择对象的父对象。
    /// </summary>
    /// <value>表示此选择对象父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 剪切当前选中的对象。
    /// </summary>
    void Cut();

    /// <summary>
    /// 复制当前选中的对象。
    /// </summary>
    void Copy();

    /// <summary>
    /// 删除当前选中的对象。
    /// </summary>
    void Delete();

    /// <summary>
    /// 取消当前选中的对象。
    /// </summary>
    void Unselect();

    /// <summary>
    /// 获取当前选中的类型。
    /// </summary>
    /// <value>表示选中类型的 <see cref="PpSelectionType"/> 枚举值。</value>
    PpSelectionType Type { get; }

    /// <summary>
    /// 获取当前选中的幻灯片范围。
    /// </summary>
    /// <value>表示选中幻灯片范围的 <see cref="IPowerPointSlideRange"/> 对象。</value>
    IPowerPointSlideRange? SlideRange { get; }

    /// <summary>
    /// 获取当前选中的形状范围。
    /// </summary>
    /// <value>表示选中形状范围的 <see cref="IPowerPointShapeRange"/> 对象。</value>
    IPowerPointShapeRange? ShapeRange { get; }

    /// <summary>
    /// 获取当前选中的文本范围。
    /// </summary>
    /// <value>表示选中文本范围的 <see cref="IPowerPointTextRange"/> 对象。</value>
    IPowerPointTextRange? TextRange { get; }

    /// <summary>
    /// 获取当前选中的子形状范围。
    /// </summary>
    /// <value>表示选中子形状范围的 <see cref="IPowerPointShapeRange"/> 对象。</value>
    IPowerPointShapeRange? ChildShapeRange { get; }

    /// <summary>
    /// 获取一个值，指示当前选择中是否包含子形状范围。
    /// </summary>
    /// <value>指示是否包含子形状范围的布尔值。</value>
    bool HasChildShapeRange { get; }

    /// <summary>
    /// 获取当前选中的文本范围 2.0 对象。
    /// </summary>
    /// <value>表示选中文本范围的 <see cref="IOfficeTextRange2"/> 对象。</value>
    IOfficeTextRange2? TextRange2 { get; }
}
