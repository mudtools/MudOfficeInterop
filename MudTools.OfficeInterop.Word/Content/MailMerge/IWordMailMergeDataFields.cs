//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示邮件合并数据源中当前活动记录的所有数据字段的集合的二次封装接口。
/// 此接口允许通过字段名称访问和修改当前记录中的具体数据值 [[1]]。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordMailMergeDataFields : IEnumerable<IWordMailMergeDataField?>, IOfficeObject<IWordMailMergeDataFields>, IDisposable
{
    /// <summary>
    /// 获取此数据字段集合所属的 Word 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此数据字段集合的父对象（通常是 <see cref="IWordMailMergeDataSource"/>）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取当前活动记录中数据字段的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取集合中指定索引处的数据字段。索引从 1 开始。
    /// </summary>
    /// <param name="index">数据字段的索引（从 1 开始）。</param>
    /// <returns>指定索引处的 <see cref="IWordMailMergeDataField"/> 对象，如果索引无效则返回 null。</returns>
    IWordMailMergeDataField? this[int index] { get; }

    /// <summary>
    /// 获取集合中具有指定名称的数据字段。
    /// </summary>
    /// <param name="fieldName">要查找的字段名称。</param>
    /// <returns>具有指定名称的 <see cref="IWordMailMergeDataField"/> 对象，如果未找到则返回 null。</returns>
    /// <exception cref="ArgumentNullException">当 <paramref name="fieldName"/> 为 null 或空时抛出。</exception>
    IWordMailMergeDataField? this[string fieldName] { get; }
}