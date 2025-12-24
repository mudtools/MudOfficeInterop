//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Office Ruler Levels 的集合接口，提供对多个 Ruler Level 对象的访问功能
/// </summary>
/// <remarks>
/// 此接口继承自 IEnumerable&lt;IOfficeRulerLevel2&gt; 和 IDisposable，
/// 支持遍历集合中的所有 Ruler Level 项并提供资源释放机制
/// </remarks>
[ComCollectionWrap(ComNamespace = "MsCore"), ItemIndex]
public interface IOfficeRulerLevels2 : IEnumerable<IOfficeRulerLevel2>, IDisposable
{
    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中项的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取指定的 Ruler Level 项
    /// </summary>
    /// <param name="index">要获取的项的从零开始的索引</param>
    /// <returns>指定索引位置的 IOfficeRulerLevel2 对象</returns>
    IOfficeRulerLevel2? this[int index] { get; }
}