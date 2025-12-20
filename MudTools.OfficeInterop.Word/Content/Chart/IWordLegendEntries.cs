//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 图表图例项集合的接口。
/// 此接口封装了图表图例中各个图例项的集合，支持通过索引或名称访问特定图例项。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordLegendEntries : IEnumerable<IWordLegendEntry?>, IDisposable
{
    /// <summary>
    /// 获取与图例项集合关联的 Word 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取图例项集合的父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取图例项集合中的图例项数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过从 1 开始的索引获取集合中的图例项。
    /// </summary>
    /// <param name="Index">图例项在集合中的位置（从 1 开始）。</param>
    /// <returns>指定索引处的图例项对象。</returns>
    IWordLegendEntry? this[int Index] { get; }

    /// <summary>
    /// 通过名称获取集合中的图例项。
    /// </summary>
    /// <param name="name">要获取的图例项的名称。</param>
    /// <returns>具有指定名称的图例项对象。</returns>
    IWordLegendEntry? this[string name] { get; }
}