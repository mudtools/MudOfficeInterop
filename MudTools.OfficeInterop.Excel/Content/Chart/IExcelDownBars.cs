//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 股票图或柱形图中“下跌柱”的封装接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.DownBars
/// 用于设置下跌柱（开盘价高于收盘价）的填充、边框、可见性等。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelDownBars : IOfficeObject<IExcelDownBars>, IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 Chart）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取下跌柱的边框格式。
    /// 返回封装后的 <see cref="IExcelBorder"/> 接口。
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取下跌柱的填充格式。
    /// 返回封装后的 <see cref="IExcelChartFillFormat"/> 接口。
    /// </summary>
    IExcelChartFillFormat? Fill { get; }

    /// <summary>
    /// 获取下跌柱的内部填充格式。
    /// 返回封装后的 <see cref="IExcelInterior"/> 接口，用于设置下跌柱的背景色、图案等内部样式属性。
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取下跌柱的图表格式对象。
    /// 返回封装后的 <see cref="IExcelChartFormat"/> 接口，用于设置下跌柱的通用格式属性，如填充、线条、阴影等。
    /// </summary>
    IExcelChartFormat? Format { get; }

    /// <summary>
    /// 选中此下跌柱（激活并高亮显示）。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除此下跌柱（将其设为不可见，并从图表中移除）。
    /// </summary>
    void Delete();
}