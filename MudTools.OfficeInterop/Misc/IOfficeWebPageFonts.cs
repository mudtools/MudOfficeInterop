//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 描述将文档保存为网页时使用的比例字体、比例字体大小、等宽字体和等宽字体大小的WebPageFont对象集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore")]
public interface IOfficeWebPageFonts : IOfficeObject<IOfficeWebPageFonts, MsCore.WebPageFonts>, IEnumerable<IOfficeWebPageFont?>, IDisposable
{
    /// <summary>
    /// 获取表示对象容器应用程序的Application对象。
    /// </summary>
    object? Application { get; }

    /// <summary>
    /// 获取一个32位整数，指示创建指定对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取指定集合中的项数。
    /// </summary>
    /// <value>集合中的项数。</value>
    int Count { get; }

    /// <summary>
    /// 根据指定的字符集从WebPageFonts集合中获取WebPageFont对象。
    /// </summary>
    /// <param name="index">所需的字符集类型。</param>
    /// <returns>指定字符集对应的WebPageFont对象。</returns>
    IOfficeWebPageFont this[MsoCharacterSet index] { get; }
}