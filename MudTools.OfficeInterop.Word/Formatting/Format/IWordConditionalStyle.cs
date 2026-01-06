//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.ConditionalStyle 的接口，用于操作条件样式。
/// 条件样式用于定义表格中特定部分（如标题行、奇数行等）的格式。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordConditionalStyle : IOfficeObject<IWordConditionalStyle, MsWord.ConditionalStyle>, IDisposable
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
    /// 获取条件样式的边框集合。
    /// </summary>
    IWordBorders? Borders { get; }

    /// <summary>
    /// 获取条件样式的底纹。
    /// </summary>
    IWordShading? Shading { get; }

    /// <summary>
    /// 获取条件样式的字体格式。
    /// </summary>
    IWordFont? Font { get; }

    /// <summary>
    /// 获取条件样式的段落格式。
    /// </summary>
    IWordParagraphFormat? ParagraphFormat { get; }

    /// <summary>
    /// 获取或设置条件样式内容下方的内边距（以磅为单位）。
    /// </summary>
    float BottomPadding { get; set; }

    /// <summary>
    /// 获取或设置条件样式内容上方的内边距（以磅为单位）。
    /// </summary>
    float TopPadding { get; set; }

    /// <summary>
    /// 获取或设置条件样式内容左侧的内边距（以磅为单位）。
    /// </summary>
    float LeftPadding { get; set; }

    /// <summary>
    /// 获取或设置条件样式内容右侧的内边距（以磅为单位）。
    /// </summary>
    float RightPadding { get; set; }
}