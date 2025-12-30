//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示下拉表单字段中的项目集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordListEntries : IEnumerable<IWordListEntry?>, IOfficeObject<IWordListEntries>, IDisposable
{
    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 <see cref="IWordApplication"/> 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取代表对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中的 <see cref="IWordListEntry"/> 对象数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引处的 <see cref="IWordListEntry"/> 对象。
    /// </summary>
    IWordListEntry? this[int index] { get; }

    /// <summary>
    /// 获取指定索引处的 <see cref="IWordListEntry"/> 对象。
    /// </summary>
    IWordListEntry? this[string name] { get; }

    /// <summary>
    /// 返回表示添加到下拉表单字段的项目 ListEntry 对象。
    /// </summary>
    /// <param name="name">必需。下拉表单字段项目的名称。</param>
    /// <param name="index">可选项。表示项目在列表中位置的数字。</param>
    /// <returns>新创建的 ListEntry 对象。</returns>
    IWordListEntry? Add(string name, int? index = null);

    /// <summary>
    /// 从下拉表单字段中删除所有项目。
    /// </summary>
    void Clear();
}