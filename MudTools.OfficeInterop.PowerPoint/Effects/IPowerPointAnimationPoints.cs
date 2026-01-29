//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

using System;
using System.Runtime.InteropServices;

/// <summary>
/// 表示 PowerPoint 动画路径中的关键点集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointAnimationPoints : IEnumerable<IPowerPointAnimationPoint?>, IDisposable
{

    /// <summary>
    /// 获取集合中的动画点数量。
    /// </summary>
    /// <value>集合中的动画点数量。</value>
    int Count { get; }

    /// <summary>
    /// 获取创建此动画点集合的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此动画点集合的父对象。
    /// </summary>
    /// <value>表示此动画点集合父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 通过索引获取集合中的指定动画点。
    /// </summary>
    /// <param name="index">要获取的动画点的索引（从1开始）。</param>
    /// <value>位于指定索引处的 <see cref="IPowerPointAnimationPoint"/> 对象。</value>
    IPowerPointAnimationPoint? this[int index] { get; }

    /// <summary>
    /// 在动画点集合中添加新动画点。
    /// </summary>
    /// <param name="index">新动画点要插入的位置索引。值为-1表示在末尾添加。</param>
    /// <returns>新添加的 <see cref="IPowerPointAnimationPoint"/> 对象。</returns>
    IPowerPointAnimationPoint? Add(int index = -1);

    /// <summary>
    /// 获取或设置一个值，指示动画路径是否平滑。
    /// </summary>
    /// <value>指示动画路径是否平滑的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Smooth { get; set; }
}