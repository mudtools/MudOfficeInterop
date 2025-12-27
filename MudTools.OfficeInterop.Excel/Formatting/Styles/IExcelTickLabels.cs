//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel TickLabels 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.TickLabels 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelTickLabels : IOfficeObject<IExcelTickLabels>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取刻度标签的父对象 (通常是 Axis)
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取刻度标签所在的 Application 对象
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    IExcelApplication? Application { get; }
    #endregion

    /// <summary>
    /// 删除刻度线标签对象。
    /// </summary>
    /// <returns>操作结果。</returns>
    object? Delete();

    /// <summary>
    /// 获取表示指定对象字体的Font对象。
    /// </summary>
    IExcelFont? Font { get; }

    /// <summary>
    /// 获取刻度线标签的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取或设置刻度线标签的数字格式代码。
    /// </summary>
    string NumberFormat { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示数字格式是否与单元格链接（当单元格中的数字格式更改时，标签中的数字格式也会更改）。
    /// </summary>
    bool NumberFormatLinked { get; set; }

    /// <summary>
    /// 获取或设置以用户语言显示的刻度线标签的数字格式代码。
    /// </summary>
    object NumberFormatLocal { get; set; }

    /// <summary>
    /// 获取或设置刻度线标签的文本方向。可以是-90到90度的整数值，或者是XlTickLabelOrientation枚举的常量之一。
    /// </summary>
    XlTickLabelOrientation Orientation { get; set; }

    /// <summary>
    /// 选择刻度线标签对象。
    /// </summary>
    /// <returns>操作结果。</returns>
    object? Select();

    /// <summary>
    /// 获取或设置刻度线标签的阅读顺序。
    /// </summary>
    int ReadingOrder { get; set; }

    /// <summary>
    /// 获取或设置一个值，表示当对象大小更改时，对象中的文本是否自动缩放字体大小。默认值为true。
    /// </summary>
    object AutoScaleFont { get; set; }

    /// <summary>
    /// 获取类别刻度线标签的级别数。
    /// </summary>
    int Depth { get; }

    /// <summary>
    /// 获取或设置标签级别之间的距离，以及第一级别与轴线之间的距离。
    /// </summary>
    int Offset { get; set; }

    /// <summary>
    /// 获取或设置指定刻度线标签的对齐方式。
    /// </summary>
    int Alignment { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示轴是否为多级别轴。
    /// </summary>
    bool MultiLevel { get; set; }

    /// <summary>
    /// 获取图表元素的格式设置对象。
    /// </summary>
    IExcelChartFormat? Format { get; }
}
