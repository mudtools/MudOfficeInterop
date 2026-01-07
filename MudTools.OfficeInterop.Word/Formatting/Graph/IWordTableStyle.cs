//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.TableStyle 的接口，用于操作表格样式。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordTableStyle : IOfficeObject<IWordTableStyle, MsWord.TableStyle>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置一个值，指示是否允许 Microsoft Word 将指定表格跨页拆分。
    /// </summary>
    bool AllowPageBreaks { get; set; }

    /// <summary>
    /// 获取或设置表示指定对象所有边框的 Borders 集合。
    /// </summary>
    IWordBorders? Borders { get; set; }

    /// <summary>
    /// 获取或设置表格样式中所有单元格内容下方添加的间距（以磅为单位）。
    /// </summary>
    float BottomPadding { get; set; }

    /// <summary>
    /// 获取或设置表格样式中所有单元格内容左侧添加的间距（以磅为单位）。
    /// </summary>
    float LeftPadding { get; set; }

    /// <summary>
    /// 获取或设置表格样式中所有单元格内容上方添加的间距（以磅为单位）。
    /// </summary>
    float TopPadding { get; set; }

    /// <summary>
    /// 获取或设置表格样式中所有单元格内容右侧添加的间距（以磅为单位）。
    /// </summary>
    float RightPadding { get; set; }

    /// <summary>
    /// 获取或设置表示指定行对齐方式的常量。
    /// </summary>
    WdRowAlignment Alignment { get; set; }

    /// <summary>
    /// 获取或设置表格样式中单元格之间的间距（以磅为单位）。
    /// </summary>
    float Spacing { get; set; }

    /// <summary>
    /// 获取或设置表示指定表格样式方向的常量。
    /// </summary>
    WdTableDirection TableDirection { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示使用指定样式格式化的表格行是否允许跨页换行。
    /// </summary>
    int AllowBreakAcrossPage { get; set; }

    /// <summary>
    /// 获取或设置指定表格样式的左缩进值（以磅为单位）。
    /// </summary>
    float LeftIndent { get; set; }

    /// <summary>
    /// 获取表示指定对象底纹格式设置的 Shading 对象。
    /// </summary>
    IWordShading? Shading { get; }

    /// <summary>
    /// 获取或设置一个整数值，表示在样式指定奇数或偶数行条纹时包含在条纹中的行数。
    /// </summary>
    int RowStripe { get; set; }

    /// <summary>
    /// 获取或设置一个整数值，表示在样式指定奇数或偶数列条纹时包含在条纹中的列数。
    /// </summary>
    int ColumnStripe { get; set; }

    /// <summary>
    /// 返回表示表格部分特殊样式格式化的 ConditionalStyle 对象。
    /// </summary>
    /// <param name="conditionCode">要应用格式化的表格区域。</param>
    /// <returns>指定表格区域的条件样式对象。</returns>
    IWordConditionalStyle? Condition(WdConditionCode conditionCode);
}