//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示不应用自动更正首字母大写的缩写。
/// <para>注：FirstLetterException 对象是 FirstLetterExceptions 集合的成员。</para>
/// <para>注：FirstLetterExceptions 集合包括所有排除的缩写。当 CorrectSentenceCaps 属性设置为 true 时，句点后面的第一个字符会自动大写。</para>
/// <para>注：在 FirstLetterExceptions 集合中的项后键入的字符不会大写。</para>
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordFirstLetterException : IOfficeObject<IWordFirstLetterException, MsWord.FirstLetterException>, IDisposable
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
    /// 获取集合中项的索引号。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置指定对象的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 删除指定的首字母大写异常。
    /// </summary>
    void Delete();
}
