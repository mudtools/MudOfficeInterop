//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示下拉列表或组合框内容控件中列表项的集合。
/// <para>注：使用 <see cref="Add(string, string, int)"/> 方法向此集合添加新项。</para>
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordContentControlListEntries : IEnumerable<IWordContentControlListEntry?>, IOfficeObject<IWordContentControlListEntries>, IDisposable
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
    /// 获取集合中的列表项数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取单个列表项。
    /// </summary>
    /// <param name="index">索引（从 1 开始）。</param>
    /// <returns>指定的列表项对象。</returns>
    IWordContentControlListEntry? this[int index] { get; }

    /// <summary>
    /// 向下拉列表或组合框内容控件添加新列表项。
    /// </summary>
    /// <param name="displayText">要在列表中显示的文本。</param>
    /// <param name="value">与列表项关联的编程值。</param>
    /// <param name="index">插入位置（从 1 开始）。如果省略或为 0，则添加到末尾。</param>
    /// <returns>新创建的列表项对象。</returns>
    IWordContentControlListEntry? Add(string displayText, string value, int index = 0);

    /// <summary>
    /// 从下拉列表或组合框内容控件中清除所有项。
    /// </summary>
    void Clear();
}