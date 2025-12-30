//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word中的自定义标签集合，提供对Word自定义标签的访问和管理功能
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordCustomLabels : IEnumerable<IWordCustomLabel?>, IOfficeObject<IWordCustomLabels>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取语言集合中的语言数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过整数索引获取自定义标签。
    /// </summary>
    /// <param name="index">自定义标签的索引位置（从0开始）</param>
    /// <returns>指定索引位置的自定义标签，如果不存在则返回null</returns>
    IWordCustomLabel? this[int index] { get; }

    /// <summary>
    /// 通过名称索引获取自定义标签。
    /// </summary>
    /// <param name="index">自定义标签的名称</param>
    /// <returns>指定名称的自定义标签，如果不存在则返回null</returns>
    IWordCustomLabel? this[string index] { get; }

    /// <summary>
    /// 添加一个新的自定义标签。
    /// </summary>
    /// <param name="Name">自定义标签的名称</param>
    /// <param name="DotMatrix">可选参数，指定是否为点阵标签，默认为null</param>
    /// <returns>新添加的自定义标签对象</returns>
    IWordCustomLabel Add(string Name, bool? DotMatrix = null);

}