//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel Borders 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Borders 的安全访问和操作
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelBorders : IEnumerable<IExcelBorder>, IDisposable
{

    /// <summary>
    /// 获取边框集合所在的父对象
    /// 对应 Borders.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取边框集合所在的Application对象
    /// 对应 Borders.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    #region 基础属性   

    /// <summary>
    /// 获取边框集合中的边框数量
    /// 对应 Borders.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定类型的边框对象
    /// </summary>
    /// <param name="borderType">边框类型</param>
    /// <returns>边框对象</returns>
    IExcelBorder? this[XlBordersIndex borderType] { get; }

    /// <summary>
    /// 获取或设置边框的线条样式。
    /// 对应 Borders.Value 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlLineStyle Value { get; set; }

    /// <summary>
    /// 获取或设置边框线条样式
    /// </summary>
    XlLineStyle LineStyle { get; set; }

    /// <summary>
    /// 获取或设置边框粗细
    /// </summary>
    XlBorderWeight Weight { get; set; }

    /// <summary>
    /// 获取或设置边框颜色
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color Color { get; set; }

    /// <summary>
    /// 获取或设置边框的颜色。
    /// 对应 Border.ColorIndex 属性
    /// </summary>
    XlColorIndex ColorIndex { get; set; }

    [ComPropertyWrap(NeedConvert = true)]
    Color ThemeColor { get; set; }

    [ComPropertyWrap(NeedConvert = true)]
    float TintAndShade { get; set; }
    #endregion
}
