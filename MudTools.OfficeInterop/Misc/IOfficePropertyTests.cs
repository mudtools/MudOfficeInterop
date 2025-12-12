//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Office 中属性测试条件集合的接口封装。
/// 该接口提供对属性测试条件集合的访问和管理。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore")]
public interface IOfficePropertyTests : IEnumerable<IOfficePropertyTest>, IDisposable
{
    /// <summary>
    /// 获取属性测试条件集合中项的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取属性测试条件（索引从 1 开始）。
    /// </summary>
    /// <param name="index">属性测试条件索引。</param>
    /// <returns>属性测试条件对象。</returns>
    IOfficePropertyTest this[int index] { get; }

    /// <summary>
    /// 添加新的属性测试条件。
    /// </summary>
    /// <param name="name">属性名称。</param>
    /// <param name="condition">比较条件。</param>
    /// <param name="value">比较值。</param>
    /// <param name="secondValue">第二个比较值（可选）。</param>
    /// <param name="connector"></param>
    /// <returns>新添加的属性测试条件对象。</returns>
    void Add(string name,
        MsoCondition condition,
        object value, object? secondValue = null,
        MsoConnector connector = MsoConnector.msoConnectorAnd);

    /// <summary>
    /// 移除指定索引的属性测试条件。
    /// </summary>
    /// <param name="index">要移除的属性测试条件索引。</param>
    void Remove(int index);
}