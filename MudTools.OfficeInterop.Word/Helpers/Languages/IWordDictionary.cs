//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示一个字典。
/// <para>注：表示自定义词典的 Dictionary 对象是 Dictionaries 集合的成员。</para>
/// <para>注：其他 Dictionary 对象由 Language 对象的属性返回；这些对象包括 ActiveSpellingDictionary、ActiveGrammarDictionary、ActiveThesaurusDictionary 和 ActiveHyphenationDictionary 属性。</para>
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordDictionary : IOfficeObject<IWordDictionary, MsWord.Dictionary>, IDisposable
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
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取或设置指定对象的语言。
    /// </summary>
    WdLanguageID LanguageID { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示指定的自定义词典是否专用于特定语言。
    /// </summary>
    bool LanguageSpecific { get; set; }

    /// <summary>
    /// 获取或设置指定对象的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取指定对象的磁盘或 Web 路径。
    /// </summary>
    string Path { get; }

    /// <summary>
    /// 获取一个值，该值指示指定的词典是否为只读。
    /// <para>注：对于 .lex 文件（内置校对词典），该属性返回 True；对于自定义拼写词典（.dic 文件），该属性返回 False。</para>
    /// </summary>
    bool ReadOnly { get; }

    /// <summary>
    /// 获取词典的类型。
    /// </summary>
    WdDictionaryType Type { get; }

    /// <summary>
    /// 删除指定的字典。
    /// </summary>
    void Delete();
}