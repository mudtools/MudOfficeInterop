//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// HangulHanjaConversionDictionaries 接口及实现类
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordHangulHanjaConversionDictionaries : IEnumerable<IWordDictionary?>, IOfficeObject<IWordHangulHanjaConversionDictionaries, MsWord.HangulHanjaConversionDictionaries>, IDisposable
{
    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 Application 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取集合中自定义词典的数量。
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
    /// 获取允许的自定义或转换词典的最大数量。
    /// <para>注：对于 Microsoft Word，此最大值为 10。</para>
    /// </summary>
    int Maximum { get; }

    /// <summary>
    /// 获取一个 32 位整数，该整数指示创建对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取或设置活动的自定义词典。
    /// </summary>
    IWordDictionary? ActiveCustomDictionary { get; set; }

    /// <summary>
    /// 获取内置的韩文/汉字转换词典。
    /// </summary>
    IWordDictionary? BuiltinDictionary { get; }

    /// <summary>
    /// 将新的自定义词典添加到集合中。
    /// </summary>
    /// <param name="FileName">新自定义词典的完整路径和文件名。</param>
    /// <returns>返回新添加的 <see cref="IWordDictionary"/> 对象。</returns>
    IWordDictionary? Add(string FileName);
}