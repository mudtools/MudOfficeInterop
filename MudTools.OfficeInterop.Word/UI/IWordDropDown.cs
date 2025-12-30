
namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中的一个下拉列表窗体域（Drop-Down Form Field）的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordDropDown : IDisposable
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
    IWordListEntries? ListEntries { get; }

    /// <summary>
    /// 获取一个值，该值指示此对象是否为有效的下拉列表窗体域。
    /// </summary>
    bool Valid { get; }
}