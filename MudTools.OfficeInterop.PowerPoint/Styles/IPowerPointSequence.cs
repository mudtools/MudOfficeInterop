//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 幻灯片中的动画序列。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointSequence : IEnumerable<IPowerPointEffect?>, IDisposable
{

    /// <summary>
    /// 获取集合中的动画效果数量。
    /// </summary>
    /// <value>集合中的动画效果数量。</value>
    int Count { get; }

    /// <summary>
    /// 获取创建此动画序列的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="Application"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此动画序列的父对象。
    /// </summary>
    /// <value>表示此动画序列父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 通过索引获取集合中的指定动画效果。
    /// </summary>
    /// <param name="index">要获取的动画效果的索引（从1开始）。</param>
    /// <value>位于指定索引处的 <see cref="IPowerPointEffect"/> 对象。</value>
    IPowerPointEffect? this[int index] { get; }

    /// <summary>
    /// 为指定形状添加动画效果。
    /// </summary>
    /// <param name="shape">要添加动画效果的形状。</param>
    /// <param name="effectId">动画效果类型。</param>
    /// <param name="level">动画效果的层级。</param>
    /// <param name="trigger">动画效果的触发类型。</param>
    /// <param name="index">新动画效果要插入的位置索引。值为-1表示在末尾添加。</param>
    /// <returns>新添加的 <see cref="IPowerPointEffect"/> 对象。</returns>
    IPowerPointEffect? AddEffect(IPowerPointShape shape, MsoAnimEffect effectId, MsoAnimateByLevel level = MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType trigger = MsoAnimTriggerType.msoAnimTriggerOnPageClick, int index = -1);

    /// <summary>
    /// 克隆指定的动画效果。
    /// </summary>
    /// <param name="effect">要克隆的动画效果。</param>
    /// <param name="index">克隆效果要插入的位置索引。值为-1表示在末尾添加。</param>
    /// <returns>克隆的 <see cref="IPowerPointEffect"/> 对象。</returns>
    IPowerPointEffect? Clone(IPowerPointEffect effect, int index = -1);

    /// <summary>
    /// 查找指定形状的第一个动画效果。
    /// </summary>
    /// <param name="shape">要查找动画效果的形状。</param>
    /// <returns>找到的 <see cref="IPowerPointEffect"/> 对象。</returns>
    IPowerPointEffect? FindFirstAnimationFor(IPowerPointShape shape);

    /// <summary>
    /// 查找指定点击次数的第一个动画效果。
    /// </summary>
    /// <param name="click">点击次数。</param>
    /// <returns>找到的 <see cref="IPowerPointEffect"/> 对象。</returns>
    IPowerPointEffect? FindFirstAnimationForClick(int click);

    /// <summary>
    /// 将动画效果转换为指定层级的构建动画。
    /// </summary>
    /// <param name="effect">要转换的动画效果。</param>
    /// <param name="level">目标层级。</param>
    /// <returns>转换后的 <see cref="IPowerPointEffect"/> 对象。</returns>
    IPowerPointEffect? ConvertToBuildLevel(IPowerPointEffect effect, MsoAnimateByLevel level);

    /// <summary>
    /// 将动画效果转换为后续效果。
    /// </summary>
    /// <param name="effect">要转换的动画效果。</param>
    /// <param name="after">后续效果类型。</param>
    /// <param name="dimColor">变暗颜色的 RGB 值。</param>
    /// <param name="dimSchemeColor">变暗颜色在颜色方案中的索引。</param>
    /// <returns>转换后的 <see cref="IPowerPointEffect"/> 对象。</returns>
    IPowerPointEffect? ConvertToAfterEffect(IPowerPointEffect effect, MsoAnimAfterEffect after, int dimColor = 0, PpColorSchemeIndex dimSchemeColor = PpColorSchemeIndex.ppNotSchemeColor);

    /// <summary>
    /// 将动画效果转换为背景动画。
    /// </summary>
    /// <param name="effect">要转换的动画效果。</param>
    /// <param name="animateBackground">指示是否为背景动画的布尔值。</param>
    /// <returns>转换后的 <see cref="IPowerPointEffect"/> 对象。</returns>
    IPowerPointEffect? ConvertToAnimateBackground(IPowerPointEffect effect, [ConvertTriState] bool animateBackground);

    /// <summary>
    /// 将动画效果转换为反向动画。
    /// </summary>
    /// <param name="effect">要转换的动画效果。</param>
    /// <param name="animateInReverse">指示是否为反向动画的布尔值。</param>
    /// <returns>转换后的 <see cref="IPowerPointEffect"/> 对象。</returns>
    IPowerPointEffect? ConvertToAnimateInReverse(IPowerPointEffect effect, [ConvertTriState] bool animateInReverse);

    /// <summary>
    /// 将动画效果转换为文本单位效果。
    /// </summary>
    /// <param name="effect">要转换的动画效果。</param>
    /// <param name="unitEffect">文本单位效果类型。</param>
    /// <returns>转换后的 <see cref="IPowerPointEffect"/> 对象。</returns>
    IPowerPointEffect? ConvertToTextUnitEffect(IPowerPointEffect effect, MsoAnimTextUnitEffect unitEffect);

    /// <summary>
    /// 添加触发式动画效果。
    /// </summary>
    /// <param name="pShape">要添加动画效果的形状。</param>
    /// <param name="effectId">动画效果类型。</param>
    /// <param name="trigger">动画效果的触发类型。</param>
    /// <param name="pTriggerShape">触发动画效果的形状。</param>
    /// <param name="bookmark">书签名称。</param>
    /// <param name="level">动画效果的层级。</param>
    /// <returns>新添加的 <see cref="IPowerPointEffect"/> 对象。</returns>
    IPowerPointEffect? AddTriggerEffect(IPowerPointShape pShape, MsoAnimEffect effectId, MsoAnimTriggerType trigger, IPowerPointShape pTriggerShape, string bookmark = "", MsoAnimateByLevel level = MsoAnimateByLevel.msoAnimateLevelNone);
}