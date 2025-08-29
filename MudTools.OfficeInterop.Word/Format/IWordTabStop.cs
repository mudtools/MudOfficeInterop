namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.TabStop 的接口，用于操作单个制表符。
/// </summary>
public interface IWordTabStop : IDisposable
{
    /// <summary>
    /// 获取或设置制表符的位置（磅）。
    /// </summary>
    float Position { get; set; }

    /// <summary>
    /// 获取或设置制表符的对齐方式。
    /// </summary>
    WdTabAlignment Alignment { get; set; }

    /// <summary>
    /// 获取或设置制表符的前导符。
    /// </summary>
    WdTabLeader Leader { get; set; }

    /// <summary>
    /// 获取制表符是否自定义。
    /// </summary>
    bool CustomTab { get; }

    /// <summary>
    /// 删除此制表符。
    /// </summary>
    void Clear();
}