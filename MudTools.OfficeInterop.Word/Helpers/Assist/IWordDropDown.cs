
namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中的一个下拉列表窗体域（Drop-Down Form Field）的封装接口。
/// </summary>
public interface IWordDropDown : IDisposable
{
    /// <summary>
    /// 获取或设置下拉列表中当前选中项的索引（从1开始）。
    /// </summary>
    int Value { get; set; }

    /// <summary>
    /// 获取或设置下拉列表的默认选中项索引（从1开始）。
    /// </summary>
    int Default { get; set; }

    /// <summary>
    /// 获取下拉列表中的所有选项项。
    /// </summary>
    IReadOnlyList<string> ListEntries { get; }

    /// <summary>
    /// 获取一个值，该值指示此对象是否为有效的下拉列表窗体域。
    /// </summary>
    bool Valid { get; }
}