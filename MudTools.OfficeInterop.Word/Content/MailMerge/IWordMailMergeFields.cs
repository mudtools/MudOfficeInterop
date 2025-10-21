//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中所有邮件合并域的集合的二次封装接口。
/// 此接口允许枚举、访问特定域，并向文档中添加新的邮件合并域 [[1]]。
/// </summary>
public interface IWordMailMergeFields : IEnumerable<IWordMailMergeField>, IDisposable
{
    /// <summary>
    /// 获取此邮件合并域集合所属的 Word 应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此邮件合并域集合的父对象（通常是 <see cref="IWordRange"/> 或 <see cref="IWordDocument"/>）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取文档中邮件合并域的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取集合中指定索引处的邮件合并域。索引从 1 开始。
    /// </summary>
    /// <param name="index">邮件合并域的索引（从 1 开始）。</param>
    /// <returns>指定索引处的 <see cref="IWordMailMergeField"/> 对象，如果索引无效则返回 null。</returns>
    IWordMailMergeField? this[int index] { get; }

    /// <summary>
    /// 在指定的范围内向文档添加一个新的邮件合并域。
    /// 新域将被插入到范围的起始位置 [[18]]。
    /// </summary>
    /// <param name="range">要插入邮件合并域的位置。</param>
    /// <param name="fieldName">要引用的数据源字段名称。</param>
    /// <returns>新创建的 <see cref="IWordMailMergeField"/> 对象。</returns>
    /// <exception cref="ArgumentNullException">当 <paramref name="range"/> 或 <paramref name="fieldName"/> 为 null 或空时抛出。</exception>
    /// <exception cref="InvalidOperationException">当添加邮件合并域操作失败时抛出。</exception>
    IWordMailMergeField Add(IWordRange range, string fieldName);
}