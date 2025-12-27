//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示所有首字母大写异常（即不自动大写的缩写）的集合。
/// <para>注：使用 FirstLetterExceptions 属性可返回 FirstLetterExceptions 集合。</para>
/// <para>注：使用 FirstLetterExceptions(index)（其中 index 是缩写或索引号）可返回单个 FirstLetterException 对象。</para>
/// </summary>
public interface IWordFirstLetterExceptions : IEnumerable<IWordFirstLetterException>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中的首字母大写异常数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引号或缩写名称获取单个首字母大写异常。
    /// </summary>
    /// <param name="index">索引号（整数）或缩写名称（字符串）。</param>
    /// <returns>指定的首字母大写异常对象。</returns>
    IWordFirstLetterException this[object index] { get; }

    /// <summary>
    /// 将缩写添加到首字母大写异常列表中。
    /// </summary>
    /// <param name="name">要添加的缩写。</param>
    /// <returns>表示添加的首字母大写异常的对象。</returns>
    IWordFirstLetterException Add(string name);
}