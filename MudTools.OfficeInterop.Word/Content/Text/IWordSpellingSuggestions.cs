//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中的拼写建议集合
/// 提供对Word文档中拼写错误的建议访问
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordSpellingSuggestions : IEnumerable<IWordSpellingSuggestion?>, IDisposable
{

    /// <summary>
    /// 获取与当前拼写建议集合关联的Word应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取当前拼写建议集合的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中拼写建议的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取拼写错误的类型
    /// </summary>
    WdSpellingErrorType SpellingErrorType { get; }

    /// <summary>
    /// 根据索引获取指定位置的拼写建议
    /// </summary>
    /// <param name="index">拼写建议在集合中的索引（从0开始）</param>
    /// <returns>指定位置的拼写建议对象，如果不存在则返回null</returns>
    IWordSpellingSuggestion? this[int index] { get; }
}