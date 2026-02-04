//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 动画效果的行为集合。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointAnimationBehaviors : IOfficeObject<IPowerPointAnimationBehaviors, MsPowerPoint.AnimationBehaviors>, IDisposable
{

    /// <summary>
    /// 获取集合中的动画行为数量。
    /// </summary>
    /// <value>集合中的动画行为数量。</value>
    int Count { get; }

    /// <summary>
    /// 获取创建此动画行为集合的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此动画行为集合的父对象。
    /// </summary>
    /// <value>表示此动画行为集合父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 通过索引获取集合中的指定动画行为。
    /// </summary>
    /// <param name="index">要获取的动画行为的索引（从1开始）。</param>
    /// <value>位于指定索引处的 <see cref="IPowerPointAnimationBehavior"/> 对象。</value>
    IPowerPointAnimationBehavior? this[int index] { get; }

    /// <summary>
    /// 在动画行为集合中添加新动画行为。
    /// </summary>
    /// <param name="type">要添加的动画行为类型。</param>
    /// <param name="index">新动画行为要插入的位置索引。值为-1表示在末尾添加。</param>
    /// <returns>新添加的 <see cref="IPowerPointAnimationBehavior"/> 对象。</returns>
    IPowerPointAnimationBehavior? Add(MsoAnimType type, int index = -1);
}