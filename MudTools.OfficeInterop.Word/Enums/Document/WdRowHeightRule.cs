namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定Word文档中表格行高度的规则
/// </summary>
public enum WdRowHeightRule
{
    /// <summary>
    /// 行高自动调整，根据内容自动确定最佳行高
    /// </summary>
    wdRowHeightAuto,
    /// <summary>
    /// 行高至少为指定值，实际行高可根据内容自动增加
    /// </summary>
    wdRowHeightAtLeast,
    /// <summary>
    /// 行高精确等于指定值，不会根据内容自动调整
    /// </summary>
    wdRowHeightExactly
}