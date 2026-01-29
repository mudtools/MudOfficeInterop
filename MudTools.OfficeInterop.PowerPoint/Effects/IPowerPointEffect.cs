//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 幻灯片中的动画效果。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointEffect : IDisposable
{
    /// <summary>
    /// 获取创建此动画效果的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此动画效果的父对象。
    /// </summary>
    /// <value>表示此动画效果父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置应用此动画效果的形状。
    /// </summary>
    /// <value>表示形状的 <see cref="IPowerPointShape"/> 对象。</value>
    IPowerPointShape? Shape { get; set; }

    /// <summary>
    /// 将此动画效果移动到指定位置。
    /// </summary>
    /// <param name="toPos">要移动到的目标位置索引。</param>
    void MoveTo(int toPos);

    /// <summary>
    /// 将此动画效果移动到指定效果之前。
    /// </summary>
    /// <param name="effect">目标效果，此效果将移动到该效果之前。</param>
    void MoveBefore(IPowerPointEffect effect);

    /// <summary>
    /// 将此动画效果移动到指定效果之后。
    /// </summary>
    /// <param name="effect">目标效果，此效果将移动到该效果之后。</param>
    void MoveAfter(IPowerPointEffect effect);

    /// <summary>
    /// 删除此动画效果。
    /// </summary>
    void Delete();

    /// <summary>
    /// 获取此动画效果在集合中的索引。
    /// </summary>
    /// <value>表示索引的整数值。</value>
    int Index { get; }

    /// <summary>
    /// 获取此动画效果的时间设置。
    /// </summary>
    /// <value>表示时间设置的 <see cref="IPowerPointTiming"/> 对象。</value>
    IPowerPointTiming? Timing { get; }

    /// <summary>
    /// 获取或设置动画效果的类型。
    /// </summary>
    /// <value>表示动画效果类型的 <see cref="MsoAnimEffect"/> 枚举值。</value>
    MsoAnimEffect EffectType { get; set; }

    /// <summary>
    /// 获取此动画效果的参数设置。
    /// </summary>
    /// <value>表示效果参数的 <see cref="IPowerPointEffectParameters"/> 对象。</value>
    IPowerPointEffectParameters? EffectParameters { get; }

    /// <summary>
    /// 获取动画效果的文本范围起始位置。
    /// </summary>
    /// <value>表示文本范围起始位置的整数值。</value>
    int TextRangeStart { get; }

    /// <summary>
    /// 获取动画效果的文本范围长度。
    /// </summary>
    /// <value>表示文本范围长度的整数值。</value>
    int TextRangeLength { get; }

    /// <summary>
    /// 获取或设置应用动画效果的段落编号。
    /// </summary>
    /// <value>表示段落编号的整数值。</value>
    int Paragraph { get; set; }

    /// <summary>
    /// 获取此动画效果的显示名称。
    /// </summary>
    /// <value>表示动画效果显示名称的字符串。</value>
    string? DisplayName { get; }

    /// <summary>
    /// 获取或设置一个值，指示此动画效果是否为退出效果。
    /// </summary>
    /// <value>指示是否为退出效果的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Exit { get; set; }

    /// <summary>
    /// 获取此动画效果的行为集合。
    /// </summary>
    /// <value>表示动画行为集合的 <see cref="IPowerPointAnimationBehaviors"/> 对象。</value>
    IPowerPointAnimationBehaviors? Behaviors { get; }

    /// <summary>
    /// 获取此动画效果的信息。
    /// </summary>
    /// <value>表示动画效果信息的 <see cref="IPowerPointEffectInformation"/> 对象。</value>
    IPowerPointEffectInformation? EffectInformation { get; }
}