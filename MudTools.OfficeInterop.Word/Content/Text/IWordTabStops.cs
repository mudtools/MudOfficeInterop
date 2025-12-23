//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.TabStops 的接口，用于操作段落的制表符集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordTabStops : IEnumerable<IWordTabStop?>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取制表符的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取制表符（从1开始）。
    /// </summary>
    IWordTabStop? this[int index] { get; }

    /// <summary>
    /// 根据位置获取制表符。
    /// </summary>
    IWordTabStop? this[float position] { get; }

    /// <summary>
    /// 添加一个新的制表符。
    /// </summary>
    /// <param name="position">制表符位置（磅）。</param>
    /// <param name="alignment">制表符对齐方式。</param>
    /// <param name="leader">制表符前导符。</param>
    /// <returns>新添加的制表符。</returns>
    IWordTabStop? Add(float position,
        WdTabAlignment alignment = WdTabAlignment.wdAlignTabLeft,
        WdTabLeader leader = WdTabLeader.wdTabLeaderSpaces);

    /// <summary>
    /// 清除所有制表符。
    /// </summary>
    void ClearAll();

    /// <summary>
    /// 获取指定位置之前的制表符。
    /// </summary>
    /// <param name="position">要查找的位置（磅）。</param>
    /// <returns>指定位置之前的制表符，如果未找到则返回null。</returns>
    IWordTabStop? Before(float position);

    /// <summary>
    /// 获取指定位置之后的制表符。
    /// </summary>
    /// <param name="position">要查找的位置（磅）。</param>
    /// <returns>指定位置之后的制表符，如果未找到则返回null。</returns>
    IWordTabStop? After(float position);
}