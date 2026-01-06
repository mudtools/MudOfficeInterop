//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示包含活动自定义拼写词典的对象集合。
/// <para>注：使用 CustomDictionaries 属性可返回当前活动的自定义字典的集合。</para>
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordDictionaries : IEnumerable<IWordDictionary?>, IOfficeObject<IWordDictionaries, MsWord.Dictionaries>, IDisposable
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
    /// 获取集合中的活动自定义词典数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引（字典名称或索引号）获取单个活动自定义词典。
    /// </summary>
    /// <param name="index">字典名称（字符串）或索引号（整数）。</param>
    /// <returns>指定的活动自定义词典对象。</returns>
    IWordDictionary? this[int index] { get; }

    /// <summary>
    /// 通过索引（字典名称或索引号）获取单个活动自定义词典。
    /// </summary>
    /// <param name="name">字典名称（字符串）或索引号（整数）。</param>
    /// <returns>指定的活动自定义词典对象。</returns>
    IWordDictionary? this[string name] { get; }

    /// <summary>
    /// 获取或设置一个 Dictionary 对象，该对象代表将向其添加单词的自定义词典。
    /// </summary>
    IWordDictionary? ActiveCustomDictionary { get; set; }

    /// <summary>
    /// 获取允许的自定义或转换词典的最大数量。
    /// <para>注：对于 Microsoft Word，此最大值为 10。</para>
    /// </summary>
    int Maximum { get; }

    /// <summary>
    /// 将新的自定义词典添加到活动自定义词典的集合。
    /// <para>注：如果没有由 FileName 指定名称的文件，Word 会创建该文件。</para>
    /// </summary>
    /// <param name="fileName">要添加的词典的完整路径和文件名。</param>
    /// <returns>表示添加的自定义词典的对象。</returns>
    IWordDictionary? Add(string fileName);

    /// <summary>
    /// 卸载所有的自定义或转换词典。
    /// <para>注：此方法并不删除词典文件，只是将它们从活动集合中移除。</para>
    /// </summary>
    void ClearAll();
}