//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Style 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Style 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelStyle : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取样式所在的父对象
    /// 对应 Style.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取样式所在的Application对象
    /// 对应 Style.Application 属性
    /// </summary>
    IExcelApplication? Application { get; }

    #endregion

    /// <summary>
    /// 获取或设置一个布尔值，表示当单元格中的文本对齐方式设置为水平或垂直均等分布时，文本是否自动缩进。
    /// </summary>
    bool AddIndent { get; set; }

    /// <summary>
    /// 获取一个布尔值，表示样式是否为内置样式。
    /// </summary>
    bool BuiltIn { get; }

    /// <summary>
    /// 获取表示样式边框的Borders集合。
    /// </summary>
    IExcelBorders? Borders { get; }

    /// <summary>
    /// 删除此样式对象。
    /// </summary>
    /// <returns>操作结果。</returns>
    object? Delete();

    /// <summary>
    /// 获取表示指定对象字体的Font对象。
    /// </summary>
    IExcelFont? Font { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示在工作表受保护时公式是否隐藏。
    /// </summary>
    bool FormulaHidden { get; set; }

    /// <summary>
    /// 获取或设置指定对象的水平对齐方式。
    /// </summary>
    XlHAlign HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示样式是否包含对齐属性（包括AddIndent、HorizontalAlignment、VerticalAlignment、WrapText和Orientation属性）。
    /// </summary>
    bool IncludeAlignment { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示样式是否包含边框属性（包括边框的颜色、颜色索引、线型和粗细属性）。
    /// </summary>
    bool IncludeBorder { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示样式是否包含字体属性（包括字体的背景、粗体、颜色、颜色索引、字体样式、斜体、名称、轮廓字体、阴影、大小、删除线、下标、上标和下划线属性）。
    /// </summary>
    bool IncludeFont { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示样式是否包含数字格式属性（NumberFormat属性）。
    /// </summary>
    bool IncludeNumber { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示样式是否包含图案属性（包括内部区域的色彩、颜色索引、负值时反转、图案、图案颜色和图案颜色索引属性）。
    /// </summary>
    bool IncludePatterns { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示样式是否包含保护属性（包括FormulaHidden和Locked属性）。
    /// </summary>
    bool IncludeProtection { get; set; }

    /// <summary>
    /// 获取或设置样式的缩进级别。
    /// </summary>
    int IndentLevel { get; set; }

    /// <summary>
    /// 获取表示指定对象内部区域的Interior对象。
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示对象是否被锁定。当工作表受保护时，如果为true则对象被锁定，如果为false则对象可以被修改。
    /// </summary>
    bool Locked { get; set; }

    /// <summary>
    /// 获取或设置一个值，表示样式是否包含合并的单元格。
    /// </summary>
    object MergeCells { get; set; }

    /// <summary>
    /// 获取样式的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取以用户语言显示的样式名称。
    /// </summary>
    string NameLocal { get; }

    /// <summary>
    /// 获取或设置对象的格式代码。
    /// </summary>
    string NumberFormat { get; set; }

    /// <summary>
    /// 获取或设置以用户语言显示的对象的格式代码。
    /// </summary>
    string NumberFormatLocal { get; set; }

    /// <summary>
    /// 获取或设置文本方向。可以是-90到90度的整数值，或者是XlOrientation枚举的常量之一。
    /// </summary>
    XlOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示文本是否自动收缩以适应可用的列宽。
    /// </summary>
    bool ShrinkToFit { get; set; }

    /// <summary>
    /// 获取指定样式的名称。
    /// </summary>
    string Value { get; }

    /// <summary>
    /// 获取或设置指定对象的垂直对齐方式。
    /// </summary>
    XlVAlign VerticalAlignment { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示Excel是否在对象中自动换行文本。
    /// </summary>
    bool WrapText { get; set; }

    /// <summary>
    /// 获取或设置指定对象的阅读顺序。
    /// </summary>
    int ReadingOrder { get; set; }
}