//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示邮件合并数据源中所有字段名称的集合的二次封装接口。
/// 此接口允许枚举和访问数据源中的各个字段。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordMailMergeFieldNames : IEnumerable<IWordMailMergeFieldName?>, IDisposable
{
    /// <summary>
    /// 获取此字段名称集合所属的 Word 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此字段名称集合的父对象（通常是 <see cref="IWordMailMergeDataSource"/>）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取数据源中字段的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取集合中指定索引处的字段名称。索引从 1 开始。
    /// </summary>
    /// <param name="index">字段名称的索引（从 1 开始）。</param>
    /// <returns>指定索引处的 <see cref="IWordMailMergeFieldName"/> 对象，如果索引无效则返回 null。</returns>
    IWordMailMergeFieldName? this[int index] { get; }

    /// <summary>
    /// 获取集合中指定索引处的字段名称。索引从 1 开始。
    /// </summary>
    /// <param name="name">字段名称的索引（从 1 开始）。</param>
    /// <returns>指定索引处的 <see cref="IWordMailMergeFieldName"/> 对象，如果索引无效则返回 null。</returns>
    IWordMailMergeFieldName? this[string name] { get; }
}