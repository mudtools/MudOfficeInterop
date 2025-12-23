//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中的智能标记集合
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordSmartTags : IEnumerable<IWordSmartTag?>, IDisposable
{
    /// <summary>
    /// 获取与智能标记集合关联的Word应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取智能标记集合的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中智能标记的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取指定位置的智能标记
    /// </summary>
    /// <param name="index">智能标记在集合中的索引（从0开始）</param>
    /// <returns>位于指定索引处的智能标记，如果不存在则返回null</returns>
    IWordSmartTag? this[int index] { get; }

    /// <summary>
    /// 通过名称获取指定的智能标记
    /// </summary>
    /// <param name="name">智能标记的名称</param>
    /// <returns>具有指定名称的智能标记，如果不存在则返回null</returns>
    IWordSmartTag? this[string name] { get; }

    /// <summary>
    /// 向集合中添加一个新的智能标记
    /// </summary>
    /// <param name="name">智能标记的名称</param>
    /// <param name="range">智能标记应用的文本范围，如果为null则使用默认范围</param>
    /// <param name="properties">自定义属性，如果为null则不设置自定义属性</param>
    /// <returns>新创建的智能标记对象，如果失败则返回null</returns>
    IWordSmartTag? Add(string name, IWordRange? range = null, IWordCustomProperties? properties = null);

    /// <summary>
    /// 根据类型名称获取智能标记的子集
    /// </summary>
    /// <param name="name">要筛选的智能标记类型名称</param>
    /// <returns>符合指定类型的智能标记集合，如果无匹配项则返回null</returns>
    IWordSmartTags? SmartTagsByType(string name);
}