//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Microsoft Word 中的语言对象，提供对 Word 语言设置和相关功能的访问。
/// 包括语言的 ID、名称、拼写检查、语法检查、词典等功能的访问接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordLanguage : IDisposable
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
    /// 获取语言的ID
    /// </summary>
    WdLanguageID ID { get; }

    /// <summary>
    /// 获取语言的名称
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取语言的本地化名称
    /// </summary>
    string NameLocal { get; }

    /// <summary>
    /// 获取或设置是否允许语法检查
    /// </summary>
    string DefaultWritingStyle { get; set; }

    /// <summary>
    /// 获取或设置是否允许拼写检查
    /// </summary>
    WdDictionaryType SpellingDictionaryType { get; set; }

    /// <summary>
    /// 获取当前语言的语法检查词典
    /// </summary>
    IWordDictionary? ActiveGrammarDictionary { get; }

    /// <summary>
    /// 获取当前语言的连字符词典
    /// </summary>
    IWordDictionary? ActiveHyphenationDictionary { get; }

    /// <summary>
    /// 获取当前语言的拼写检查词典
    /// </summary>
    IWordDictionary? ActiveSpellingDictionary { get; }

    /// <summary>
    /// 获取当前语言的同义词词典
    /// </summary>
    IWordDictionary? ActiveThesaurusDictionary { get; }

    /// <summary>
    /// 获取当前语言的写作风格列表
    /// </summary>
    object WritingStyleList { get; }
}