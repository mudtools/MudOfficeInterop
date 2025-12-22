//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 中“自动图文集条目”集合的封装接口。
/// 提供对模板中所有自动图文集条目的访问、添加和枚举功能。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordAutoTextEntries : IEnumerable<IWordAutoTextEntry?>, IDisposable
{
    /// <summary>
    /// 获取此集合所属的 Word 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此集合的父对象（通常是 <see cref="IWordMailMerge"/>）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中自动图文集条目的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取指定的自动图文集条目（从 1 开始）
    /// </summary>
    /// <param name="index">索引（1-based）</param>
    /// <returns>封装后的条目对象</returns>
    IWordAutoTextEntry? this[int index] { get; }

    /// <summary>
    /// 根据名称获取指定的自动图文集条目
    /// </summary>
    /// <param name="name">条目名称</param>
    /// <returns>封装后的条目对象；若不存在返回 null</returns>
    IWordAutoTextEntry? this[string name] { get; }

    /// <summary>
    /// 向模板中添加一个新的自动图文集条目
    /// </summary>
    /// <param name="name">条目名称</param>
    /// <param name="range">要插入的内容文本</param>
    /// <returns>新创建的条目封装对象</returns>
    IWordAutoTextEntry? Add(string name, IWordRange range);

    /// <summary>
    /// 将指定范围的内容追加到 Spike（Word 中的一个特殊剪贴板功能）
    /// Spike 是 Word 中一个特殊的文档片段收集器，可以收集多个文本片段
    /// </summary>
    /// <param name="range">要追加到 Spike 的内容范围</param>
    /// <returns>表示追加操作结果的自动图文集条目对象；若操作失败则返回 null</returns>
    IWordAutoTextEntry? AppendToSpike(IWordRange range);
}