//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示Office HTML项目项的集合接口，提供对集合中项目项的访问和管理功能
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore"), ItemIndex]
public interface IOfficeHTMLProjectItems : IOfficeObject<IOfficeHTMLProjectItems>, IEnumerable<IOfficeHTMLProjectItem?>, IDisposable
{
    /// <summary>
    /// 获取此集合的父对象
    /// </summary>
    /// <value>父对象，可能为 null</value>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中项目项的数量
    /// </summary>
    /// <value>集合中项目项的总数</value>
    int Count { get; }

    /// <summary>
    /// 根据索引获取集合中的项目项
    /// </summary>
    /// <param name="index">要获取的项目项的从零开始的索引</param>
    /// <returns>指定索引处的项目项；如果索引超出范围，则可能为 null</returns>
    IOfficeHTMLProjectItem? this[int index] { get; }

    /// <summary>
    /// 根据名称获取集合中的项目项
    /// </summary>
    /// <param name="name">要获取的项目项的名称</param>
    /// <returns>指定名称的项目项；如果不存在，则可能为 null</returns>
    IOfficeHTMLProjectItem? this[string name] { get; }
}