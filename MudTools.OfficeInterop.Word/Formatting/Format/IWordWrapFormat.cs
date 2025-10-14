
namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 定义对 Microsoft.Office.Interop.Word.WrapFormat 对象的二次封装接口。
/// 代表形状或图片的环绕格式设置。
/// </summary>
public interface IWordWrapFormat : IDisposable
{
    /// <summary>
    /// 获取创建 WrapFormat 对象的应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取 WrapFormat 对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置环绕类型（例如，四周型环绕、紧密型环绕、无环绕等）。
    /// </summary>
    WdWrapType Type { get; set; }

    /// <summary>
    /// 获取或设置形状或图片与环绕文本之间的距离。
    /// </summary>
    float DistanceTop { get; set; }

    /// <summary>
    /// 获取或设置形状或图片与环绕文本之间的距离。
    /// </summary>
    float DistanceBottom { get; set; }

    /// <summary>
    /// 获取或设置形状或图片与环绕文本之间的距离。
    /// </summary>
    float DistanceLeft { get; set; }

    /// <summary>
    /// 获取或设置形状或图片与环绕文本之间的距离。
    /// </summary>
    float DistanceRight { get; set; }
}