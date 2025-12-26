//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示邮件合并中所有标准字段与数据源字段之间映射关系的集合的二次封装接口。
/// 此集合包含了 Word 支持的所有预定义标准字段（如姓名、地址、城市等）的映射 [[1]]。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordMappedDataFields : IEnumerable<IWordMappedDataField?>, IOfficeObject<IWordMappedDataFields>, IDisposable
{
    /// <summary>
    /// 获取此映射字段集合所属的 Word 应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此映射字段集合的父对象（通常是 <see cref="IWordMailMerge"/>）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中映射字段的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取集合中指定索引处的映射字段。索引从 1 开始。
    /// </summary>
    /// <param name="index">映射字段的索引（从 1 开始）。</param>
    /// <returns>指定索引处的 <see cref="IWordMappedDataField"/> 对象，如果索引无效则返回 null。</returns>
    IWordMappedDataField? this[WdMappedDataFields index] { get; }
}