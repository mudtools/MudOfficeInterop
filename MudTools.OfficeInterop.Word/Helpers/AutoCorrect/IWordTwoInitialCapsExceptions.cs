//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示自动更正初始大写异常的对象集合。
/// <para>注：此集合对应于“自动更正例外项”对话框中“INitial CAps”选项卡上列出的所有项。</para>
/// <para>注：使用 TwoInitialCapsExceptions 属性可返回 TwoInitialCapsExceptions 集合。</para>
/// <para>注：使用 TwoInitialCapsExceptions(index)（其中 index 是初始大写字母名称或索引号）可返回单个 TwoInitialCapsException 对象。</para>
/// </summary>
public interface IWordTwoInitialCapsExceptions : IEnumerable<IWordTwoInitialCapsException>, IDisposable
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
    /// 获取集合中的初始大写自动更正异常数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引号或异常名称获取单个初始大写自动更正异常。
    /// </summary>
    /// <param name="index">索引号（整数）或异常名称（字符串）。</param>
    /// <returns>指定的初始大写自动更正异常对象。</returns>
    IWordTwoInitialCapsException this[object index] { get; }

    /// <summary>
    /// 将项添加到初始大写自动更正异常列表中。
    /// </summary>
    /// <param name="name">要添加的异常名称。</param>
    /// <returns>表示添加的异常的对象。</returns>
    IWordTwoInitialCapsException Add(string name);
}