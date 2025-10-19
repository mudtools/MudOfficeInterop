namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定可以编辑文档特定区域的用户类型
/// </summary>
public enum WdEditorType
{
    /// <summary>
    /// 表示所有用户都可以编辑的区域
    /// </summary>
    wdEditorEveryone = -1,

    /// <summary>
    /// 表示只有文档所有者可以编辑的区域
    /// </summary>
    wdEditorOwners = -4,

    /// <summary>
    /// 表示特定的编辑者可以编辑的区域
    /// </summary>
    wdEditorEditors = -5,

    /// <summary>
    /// 表示当前用户可以编辑的区域
    /// </summary>
    wdEditorCurrent = -6
}