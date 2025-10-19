
namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中的一个复选框窗体域（Check Box Form Field）的封装接口。
/// </summary>
public interface IWordCheckBox : IDisposable
{
    /// <summary>
    /// 获取或设置复选框是否被选中。
    /// </summary>
    bool Value { get; set; }

    /// <summary>
    /// 获取或设置复选框的默认状态（选中或未选中）。
    /// </summary>
    bool Default { get; set; }

    /// <summary>
    /// 获取或设置复选框是否根据周围文字的字号自动调整大小。
    /// </summary>
    bool AutoSize { get; set; }

    /// <summary>
    /// 获取一个值，该值指示此对象是否为有效的复选框窗体域。
    /// </summary>
    bool Valid { get; }
}