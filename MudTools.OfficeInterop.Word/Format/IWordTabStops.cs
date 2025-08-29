namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.TabStops 的接口，用于操作段落的制表符集合。
/// </summary>
public interface IWordTabStops : IEnumerable<IWordTabStop>, IDisposable
{
    /// <summary>
    /// 获取制表符的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取制表符（从1开始）。
    /// </summary>
    IWordTabStop this[int index] { get; }

    /// <summary>
    /// 根据位置获取制表符。
    /// </summary>
    IWordTabStop this[float position] { get; }

    /// <summary>
    /// 添加一个新的制表符。
    /// </summary>
    /// <param name="position">制表符位置（磅）。</param>
    /// <param name="alignment">制表符对齐方式。</param>
    /// <param name="leader">制表符前导符。</param>
    /// <returns>新添加的制表符。</returns>
    IWordTabStop Add(float position,
        WdTabAlignment alignment = WdTabAlignment.wdAlignTabLeft,
        WdTabLeader leader = WdTabLeader.wdTabLeaderSpaces);

    /// <summary>
    /// 根据位置查找制表符。
    /// </summary>
    /// <param name="position">制表符位置。</param>
    /// <returns>找到的制表符，如果不存在则返回null。</returns>
    IWordTabStop Find(float position);

    /// <summary>
    /// 清除指定位置的制表符。
    /// </summary>
    /// <param name="position">制表符位置。</param>
    void Clear(float position);

    /// <summary>
    /// 清除所有制表符。
    /// </summary>
    void ClearAll();

    /// <summary>
    /// 获取所有制表符位置的列表。
    /// </summary>
    /// <returns>制表符位置列表。</returns>
    List<float> GetPositions();
}