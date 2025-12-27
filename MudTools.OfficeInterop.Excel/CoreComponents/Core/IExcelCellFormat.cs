//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel CellFormat 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.CellFormat 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelCellFormat : IOfficeObject<IExcelCellFormat>, IDisposable
{
    /// <summary>
    /// 获取单元格格式对象的父对象（通常是 Application）
    /// 对应 CellFormat.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取单元格格式对象所在的Application对象
    /// 对应 CellFormat.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置基于单元格边框格式的搜索条件。
    /// </summary>
    IExcelBorders? Borders { get; set; }

    /// <summary>
    /// 获取或设置基于单元格字体格式的搜索条件。
    /// </summary>
    IExcelFont? Font { get; set; }

    /// <summary>
    /// 获取或设置基于单元格内部区域格式的搜索条件。
    /// </summary>
    IExcelInterior? Interior { get; set; }

    /// <summary>
    /// 获取或设置单元格的数字格式代码。如果指定范围内的所有单元格没有相同的数字格式，则返回null。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string NumberFormat { get; set; }

    /// <summary>
    /// 获取或设置以用户语言显示的单元格数字格式代码。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string NumberFormatLocal { get; set; }

    /// <summary>
    /// 获取或设置一个值，表示当单元格中的文本对齐方式设置为水平或垂直均等分布时，文本是否自动缩进。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool AddIndent { get; set; }

    /// <summary>
    /// 获取或设置单元格或区域的缩进级别。可以是0到15的整数。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    int IndentLevel { get; set; }

    /// <summary>
    /// 获取或设置指定对象的水平对齐方式。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlHAlign HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取或设置指定对象的垂直对齐方式。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlVAlign VerticalAlignment { get; set; }

    /// <summary>
    /// 获取或设置文本方向。可以是-90到90度的整数值。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置一个值，表示文本是否自动收缩以适应可用的列宽。如果此属性在指定范围内的所有单元格中未设置为相同的值，则返回null。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool ShrinkToFit { get; set; }

    /// <summary>
    /// 获取或设置一个值，表示Excel是否在对象中自动换行文本。如果指定范围包含一些自动换行文本的单元格和一些不自动换行文本的单元格，则返回null。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool WrapText { get; set; }

    /// <summary>
    /// 获取或设置一个值，表示对象是否被锁定。当工作表受保护时，如果为true则对象被锁定，如果为false则对象可以被修改。如果指定范围同时包含锁定和未锁定的单元格，则返回null。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Locked { get; set; }

    /// <summary>
    /// 获取或设置一个值，表示在工作表受保护时公式是否隐藏。如果指定范围包含一些FormulaHidden为true的单元格和一些FormulaHidden为false的单元格，则返回null。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool FormulaHidden { get; set; }

    /// <summary>
    /// 获取或设置一个值，表示范围或样式是否包含合并的单元格。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool MergeCells { get; set; }

    /// <summary>
    /// 清除在Application.FindFormat和Application.ReplaceFormat属性中设置的搜索条件。
    /// </summary>
    void Clear();
}
