//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档窗口中窗格的集合。
/// 封装了 Microsoft.Office.Interop.Word.Panes 对象。
/// </summary>
public interface IWordPanes : IEnumerable<IWordPane>, IDisposable
{
    #region 属性

    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 <see cref="IWordApplication"/> 对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取集合中窗格的计数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取代表 <see cref="IWordPanes"/> 对象的父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 通过索引（从 1 开始）获取集合中的单个窗格。
    /// </summary>
    /// <param name="index">窗格的索引号。</param>
    /// <returns>指定索引处的 <see cref="IWordPane"/> 对象。</returns>
    IWordPane this[int index] { get; }

    #endregion // 属性
}