//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel PageSetup 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.PageSetup 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelPageSetup : IDisposable
{
    /// <summary>
    /// 获取表示Excel应用程序的Application对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取指定对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示文档元素是否以黑白方式打印。
    /// </summary>
    bool BlackAndWhite { get; set; }

    /// <summary>
    /// 获取或设置底边距的大小（以磅为单位）。
    /// </summary>
    double BottomMargin { get; set; }

    /// <summary>
    /// 获取或设置页脚的中间部分。
    /// </summary>
    string CenterFooter { get; set; }

    /// <summary>
    /// 获取或设置页眉的中间部分。
    /// </summary>
    string CenterHeader { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示打印时工作表是否在页面上水平居中。
    /// </summary>
    bool CenterHorizontally { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示打印时工作表是否在页面上垂直居中。
    /// </summary>
    bool CenterVertically { get; set; }

    /// <summary>
    /// 获取或设置图表缩放以适应页面的方式。
    /// </summary>
    XlObjectSize ChartSize { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示工作表是否在没有图形的情况下打印。
    /// </summary>
    bool Draft { get; set; }

    /// <summary>
    /// 获取或设置打印此工作表时使用的起始页码。如果为xlAutomatic，则Excel自动选择起始页码。默认为xlAutomatic。
    /// </summary>
    int FirstPageNumber { get; set; }

    /// <summary>
    /// 获取或设置打印时工作表缩放的高度页数。仅适用于工作表。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    int FitToPagesTall { get; set; }

    /// <summary>
    /// 获取或设置打印时工作表缩放的宽度页数。仅适用于工作表。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    int FitToPagesWide { get; set; }

    /// <summary>
    /// 获取或设置从页面底部到页脚的距离（以磅为单位）。
    /// </summary>
    double FooterMargin { get; set; }

    /// <summary>
    /// 获取或设置从页面顶部到页眉的距离（以磅为单位）。
    /// </summary>
    double HeaderMargin { get; set; }

    /// <summary>
    /// 获取或设置页脚的左侧部分。
    /// </summary>
    string LeftFooter { get; set; }

    /// <summary>
    /// 获取或设置页眉的左侧部分。
    /// </summary>
    string LeftHeader { get; set; }

    /// <summary>
    /// 获取或设置左边距的大小（以磅为单位）。
    /// </summary>
    double LeftMargin { get; set; }

    /// <summary>
    /// 获取或设置打印大型工作表时Excel使用的页码顺序。
    /// </summary>
    XlOrder Order { get; set; }

    /// <summary>
    /// 获取或设置打印方向：纵向或横向。
    /// </summary>
    XlPageOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置纸张大小。
    /// </summary>
    XlPaperSize PaperSize { get; set; }

    /// <summary>
    /// 获取或设置要打印的区域，使用宏语言中的A1样式引用字符串表示。
    /// </summary>
    string PrintArea { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否在页面上打印单元格网格线。仅适用于工作表。
    /// </summary>
    bool PrintGridlines { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否在打印时包含行号和列标。仅适用于工作表。
    /// </summary>
    bool PrintHeadings { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否将单元格批注作为尾注与工作表一起打印。仅适用于工作表。
    /// </summary>
    bool PrintNotes { get; set; }

    /// <summary>
    /// 获取或设置要在每页左侧重复的单元格所在的列，使用宏语言中的A1样式表示法字符串。
    /// </summary>
    string PrintTitleColumns { get; set; }

    /// <summary>
    /// 获取或设置要在每页顶部重复的单元格所在的行，使用宏语言中的A1样式表示法字符串。
    /// </summary>
    string PrintTitleRows { get; set; }

    /// <summary>
    /// 获取或设置页脚的右侧部分。
    /// </summary>
    string RightFooter { get; set; }

    /// <summary>
    /// 获取或设置页眉的右侧部分。
    /// </summary>
    string RightHeader { get; set; }

    /// <summary>
    /// 获取或设置右边距的大小（以磅为单位）。
    /// </summary>
    double RightMargin { get; set; }

    /// <summary>
    /// 获取或设置上边距的大小（以磅为单位）。
    /// </summary>
    double TopMargin { get; set; }

    /// <summary>
    /// 获取或设置Excel缩放工作表以进行打印的百分比（介于10%到400%之间）。仅适用于工作表。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    int Zoom { get; set; }

    /// <summary>
    /// 获取或设置批注与工作表一起打印的方式。
    /// </summary>
    XlPrintLocation PrintComments { get; set; }

    /// <summary>
    /// 获取或设置一个XlPrintErrors常量，指定显示的打印错误类型。此功能允许用户在打印工作表时抑制错误值的显示。
    /// </summary>
    XlPrintErrors PrintErrors { get; set; }

    /// <summary>
    /// 获取表示页眉中间部分图片的Graphic对象。用于设置图片属性。
    /// </summary>
    IExcelGraphic? CenterHeaderPicture { get; }

    /// <summary>
    /// 获取表示页脚中间部分图片的Graphic对象。用于设置图片属性。
    /// </summary>
    IExcelGraphic? CenterFooterPicture { get; }

    /// <summary>
    /// 获取表示页眉左侧部分图片的Graphic对象。用于设置图片属性。
    /// </summary>
    IExcelGraphic? LeftHeaderPicture { get; }

    /// <summary>
    /// 获取表示页脚左侧部分图片的Graphic对象。用于设置图片属性。
    /// </summary>
    IExcelGraphic? LeftFooterPicture { get; }

    /// <summary>
    /// 获取表示页眉右侧部分图片的Graphic对象。用于设置图片属性。
    /// </summary>
    IExcelGraphic? RightHeaderPicture { get; }

    /// <summary>
    /// 获取表示页脚右侧部分图片的Graphic对象。用于设置图片属性。
    /// </summary>
    IExcelGraphic? RightFooterPicture { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示指定的PageSetup对象是否为奇数页和偶数页使用不同的页眉和页脚。
    /// </summary>
    bool OddAndEvenPagesHeaderFooter { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示第一页是否使用不同的页眉或页脚。
    /// </summary>
    bool DifferentFirstPageHeaderFooter { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示当文档大小更改时，页眉和页脚是否应随文档缩放。
    /// </summary>
    bool ScaleWithDocHeaderFooter { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示Excel是否将页眉和页脚与页面设置选项中设置的边距对齐。
    /// </summary>
    bool AlignMarginsHeaderFooter { get; set; }

    /// <summary>
    /// 获取或设置Pages集合中的页面计数或项目编号。
    /// </summary>
    IExcelPages? Pages { get; }

    /// <summary>
    /// 获取或设置工作簿或节中偶数页的文本对齐方式。
    /// </summary>
    IExcelPage? EvenPage { get; }

    /// <summary>
    /// 获取或设置工作簿或节中第一页的文本对齐方式。
    /// </summary>
    IExcelPage? FirstPage { get; }

}