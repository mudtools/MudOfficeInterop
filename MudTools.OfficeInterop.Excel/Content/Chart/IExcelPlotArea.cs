//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel PlotArea 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.PlotArea 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelPlotArea : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取或设置绘图区的名称
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取绘图区的父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取绘图区所在的 Application 对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }
    #endregion

    #region 位置和大小
    /// <summary>
    /// 获取或设置绘图区的左边距
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取或设置绘图区的顶边距
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取或设置绘图区的宽度
    /// </summary>
    double Width { get; set; }

    /// <summary>
    /// 获取或设置绘图区的高度
    /// </summary>
    double Height { get; set; }

    /// <summary>
    /// 获取或设置绘图区的内部左边距
    /// </summary>
    double InsideLeft { get; set; }

    /// <summary>
    /// 获取或设置绘图区的内部顶边距
    /// </summary>
    double InsideTop { get; set; }

    /// <summary>
    /// 获取或设置绘图区的内部宽度
    /// </summary>
    double InsideWidth { get; set; }

    /// <summary>
    /// 获取或设置绘图区的内部高度
    /// </summary>
    double InsideHeight { get; set; }

    XlChartElementPosition Position { get; set; }
    #endregion

    #region 格式设置
    /// <summary>
    /// 获取绘图区的字体对象
    /// </summary>
    IExcelChartFormat? Format { get; }

    /// <summary>
    /// 获取绘图区的背景填充对象
    /// </summary>
    IExcelChartFillFormat? Fill { get; }

    /// <summary>
    /// 获取绘图区的边框对象
    /// </summary>
    IExcelBorder? Border { get; }

    IExcelInterior? Interior { get; }
    #endregion

    #region 操作方法  
    /// <summary>
    /// 清除绘图区格式
    /// </summary>
    object? ClearFormats();

    object? Select();

    #endregion
}
