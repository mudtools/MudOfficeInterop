//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示关联 Range 对象的显示设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelDisplayFormat : IOfficeObject<IExcelDisplayFormat, MsExcel.DisplayFormat>, IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（如 FillFormat、GlowFormat、ShadowFormat 等）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取一个 Borders 对象，表示关联 Range 对象在当前用户界面中显示的边框。
    /// </summary>
    IExcelBorders? Borders { get; }

    /// <summary>
    /// 获取一个 Characters 对象，表示关联 Range 对象文本中的一系列字符，该对象在当前用户界面中显示。
    /// </summary>
    /// <param name="start">要返回的第一个字符。如果此参数为 1 或省略，则此属性返回从第一个字符开始的字符范围。</param>
    /// <param name="length">要返回的字符数。如果省略，则此属性返回字符串的其余部分（起始字符之后的所有内容）。</param>
    /// <returns>表示字符范围的 Characters 对象。</returns>
    [MethodIndex]
    IExcelCharacters? Characters(int? start = null, int? length = null);

    /// <summary>
    /// 获取一个 Font 对象，表示关联 Range 对象在当前用户界面中显示的字体。
    /// </summary>
    IExcelFont? Font { get; }

    /// <summary>
    /// 获取一个包含 Style 对象的值，表示关联 Range 对象在当前用户界面中显示的样式。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelStyle? Style { get; }

    /// <summary>
    /// 获取一个值，指示当单元格文本对齐方式设置为水平或垂直均匀分布时，Microsoft Excel 是否自动缩进关联 Range 对象的文本，该对象在当前用户界面中显示。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool AddIndent { get; }

    /// <summary>
    /// 获取一个值，指示在工作表受保护时，关联 Range 对象的公式是否隐藏，该对象在当前用户界面中显示。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool FormulaHidden { get; }

    /// <summary>
    /// 获取一个值，表示关联 Range 对象在当前用户界面中显示的水平对齐方式。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlHAlign HorizontalAlignment { get; }

    /// <summary>
    /// 获取一个值，表示关联 Range 对象在当前用户界面中显示的缩进级别。
    /// </summary>
    object IndentLevel { get; }

    /// <summary>
    /// 获取一个 Interior 对象，表示关联 Range 对象在当前用户界面中显示的内部区域。
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取一个值，指示关联 Range 对象在当前用户界面中显示时是否已锁定。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Locked { get; }

    /// <summary>
    /// 获取一个值，指示关联 Range 对象在当前用户界面中显示时是否包含合并单元格。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool MergeCells { get; }

    /// <summary>
    /// 获取一个值，表示关联 Range 对象在当前用户界面中显示的格式代码。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string NumberFormat { get; }

    /// <summary>
    /// 获取一个值，表示关联 Range 对象在当前用户界面中显示的格式代码（以用户语言字符串表示）。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string NumberFormatLocal { get; }

    /// <summary>
    /// 获取一个值，表示关联 Range 对象在当前用户界面中显示的文本方向。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlOrientation Orientation { get; }

    /// <summary>
    /// 获取关联 Range 对象在当前用户界面中显示的阅读顺序。
    /// </summary>
    int ReadingOrder { get; }

    /// <summary>
    /// 获取一个值，指示 Microsoft Excel 是否自动收缩文本以适应关联 Range 对象在当前用户界面中显示的可用列宽。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool ShrinkToFit { get; }

    /// <summary>
    /// 获取一个值，表示关联 Range 对象在当前用户界面中显示的垂直对齐方式。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlVAlign VerticalAlignment { get; }

    /// <summary>
    /// 获取一个值，指示 Microsoft Excel 是否自动换行关联 Range 对象的文本，该对象在当前用户界面中显示。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool WrapText { get; }
}