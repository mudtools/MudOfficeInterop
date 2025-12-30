//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Microsoft Word 中的 Excel Rectangle 集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordRectangles : IDisposable, IEnumerable<IWordRectangle?>, IOfficeObject<IWordRectangles>
{
    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 <see cref="IWordApplication"/> 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取代表对象的父对象。
    /// </summary>
    object? Parent { get; }
    /// <summary>
    /// 获取指定集合中的对象数目。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引处的对象。
    /// </summary>
    /// <param name="index"></param>
    /// <returns></returns>
    IWordRectangle? this[int index] { get; }
}