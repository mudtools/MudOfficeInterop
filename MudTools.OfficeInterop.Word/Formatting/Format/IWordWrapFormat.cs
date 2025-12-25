//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 定义对 Microsoft.Office.Interop.Word.WrapFormat 对象的二次封装接口。
/// 代表形状或图片的环绕格式设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordWrapFormat : IDisposable
{
    /// <summary>
    /// 获取创建 WrapFormat 对象的应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
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

    /// <summary>
    /// 返回或设置一个值，指定给定的形状是否可以与其他图形重叠。 可以设置为 True 或 False 。
    /// </summary>
    int AllowOverlap { get; set; }
}