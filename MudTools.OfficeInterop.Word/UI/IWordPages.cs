//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中页面布局信息的集合。
/// 封装了 Microsoft.Office.Interop.Word.Pages 对象。
/// </summary>
/// <remarks>
/// Pages 集合包含文档中每个页面的 Page 对象。
/// 此集合通常在文档处于页面布局视图或打印预览时可用。
/// </remarks>
public interface IWordPages : IEnumerable<IWordPage>, IDisposable
{
    #region 属性

    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 <see cref="IWordApplication"/> 对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取集合中页面的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取代表 <see cref="IWordPages"/> 对象的父对象（通常是 Range 对象）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 通过索引（从 1 开始）获取集合中的单个页面。
    /// </summary>
    /// <param name="index">页面的索引号。</param>
    /// <returns>指定索引处的 <see cref="IWordPage"/> 对象。</returns>
    /// <exception cref="ArgumentOutOfRangeException">如果索引超出范围。</exception>
    IWordPage this[int index] { get; }
    #endregion 

}