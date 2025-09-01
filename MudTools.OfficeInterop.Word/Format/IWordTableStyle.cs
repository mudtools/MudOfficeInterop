//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Interop.Word;

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.TableStyle 的接口，用于操作表格样式。
/// </summary>
public interface IWordTableStyle : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    WdTableDirection TableDirection { get; set; }
    /// <summary>
    /// 获取或设置表格左边框样式。
    /// </summary>
    IWordBorder? LeftBorder { get; }

    /// <summary>
    /// 获取或设置表格右边框样式。
    /// </summary>
    IWordBorder? RightBorder { get; }

    /// <summary>
    /// 获取或设置表格上边框样式。
    /// </summary>
    IWordBorder? TopBorder { get; }

    /// <summary>
    /// 获取或设置表格下边框样式。
    /// </summary>
    IWordBorder? BottomBorder { get; }

    /// <summary>
    /// 获取或设置表格水平边框样式。
    /// </summary>
    IWordBorder? HorizontalBorder { get; }

    /// <summary>
    /// 获取或设置表格垂直边框样式。
    /// </summary>
    IWordBorder? VerticalBorder { get; }

    bool AllowBreakAcrossPage { get; set; }

    /// <summary>
    /// 获取或设置是否允许跨页断行。
    /// </summary>
    bool AllowPageBreaks { get; set; }

    /// <summary>
    /// 获取或设置表格行的底纹。
    /// </summary>
    IWordShading Shading { get; }

    /// <summary>
    /// 获取或设置表格的对齐方式。
    /// </summary>
    WdRowAlignment Alignment { get; set; }

    int ColumnStripe { get; set; }

    int RowStripe { get; set; }

    /// <summary>
    /// 获取或设置表格缩进距离。
    /// </summary>
    float LeftIndent { get; set; }

    /// <summary>
    /// 获取或设置表格左边距。
    /// </summary>
    float LeftPadding { get; set; }

    /// <summary>
    /// 获取或设置表格右边距。
    /// </summary>
    float RightPadding { get; set; }

    /// <summary>
    /// 获取或设置表格上边距。
    /// </summary>
    float TopPadding { get; set; }

    /// <summary>
    /// 获取或设置表格下边距。
    /// </summary>
    float BottomPadding { get; set; }

    float Spacing { get; set; }

    IWordConditionalStyle Condition(WdConditionCode conditionCode);
}